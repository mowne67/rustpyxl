use crate::{Cell, Row, read, write};
use pyo3::{
    exceptions::{PyTypeError, PyValueError},
    prelude::*,
    types::{PyAny, PyBool, PyBytes, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple},
};
use std::{
    collections::VecDeque,
    sync::{
        Arc, Mutex,
        mpsc::{Receiver, sync_channel},
    },
};

fn error(e: impl std::fmt::Display) -> PyErr {
    PyValueError::new_err(e.to_string())
}

#[pyclass]
struct NativeWorkbook {
    workbook: Arc<read::Workbook>,
}

#[pymethods]
impl NativeWorkbook {
    #[new]
    fn new(py: Python<'_>, path: String) -> PyResult<Self> {
        let workbook = py.detach(|| read::Workbook::open(path)).map_err(error)?;
        Ok(Self {
            workbook: Arc::new(workbook),
        })
    }

    #[getter]
    fn sheetnames(&self) -> Vec<String> {
        self.workbook
            .sheets
            .iter()
            .map(|s| s.name.clone())
            .collect()
    }

    #[getter]
    fn shared_strings_path(&self) -> Option<String> {
        self.workbook.shared_strings_path.clone()
    }

    fn shared_len(&self) -> usize {
        self.workbook.shared_len()
    }

    fn shared_string(&self, index: usize) -> PyResult<String> {
        self.workbook
            .shared_string(index)
            .map(str::to_owned)
            .ok_or_else(|| {
                pyo3::exceptions::PyIndexError::new_err("shared string index out of range")
            })
    }

    fn rows(&self, py: Python<'_>, sheet: usize, data_only: bool) -> PyResult<NativeRows> {
        let mut rows = py
            .detach(|| self.workbook.rows(sheet, data_only))
            .map_err(error)?;
        // Parse ahead on a worker thread so XML decoding overlaps Python's use
        // of the previous batch. The bounded channel caps read-ahead memory;
        // dropping the receiver stops the worker after its current batch.
        let (sender, receiver) = sync_channel(2);
        std::thread::spawn(move || {
            loop {
                let batch = rows.batch(1024);
                let done = batch.as_ref().map_or(true, Vec::is_empty);
                if sender.send(batch).is_err() || done {
                    break;
                }
            }
        });
        Ok(NativeRows {
            rows: Mutex::new(Some(Prefetch {
                receiver: Some(receiver),
                ready: VecDeque::new(),
            })),
        })
    }
}

struct Prefetch {
    receiver: Option<Receiver<crate::Result<Vec<Row>>>>,
    ready: VecDeque<Row>,
}

#[pyclass(weakref)]
struct NativeRows {
    rows: Mutex<Option<Prefetch>>,
}

#[pymethods]
impl NativeRows {
    fn batch(&self, py: Python<'_>, size: usize) -> PyResult<Py<PyList>> {
        if !(1..=65_536).contains(&size) {
            return Err(error("batch size must be in 1..65536"));
        }
        let rows = py
            .detach(|| -> crate::Result<Vec<Row>> {
                let mut guard = self.rows.lock().map_err(|_| "reader lock poisoned")?;
                let state = guard.as_mut().ok_or("iterator is closed")?;
                while state.ready.len() < size {
                    let Some(receiver) = &state.receiver else {
                        break;
                    };
                    match receiver.recv() {
                        Ok(Ok(batch)) if !batch.is_empty() => state.ready.extend(batch),
                        Ok(Err(e)) => {
                            // Undelivered rows are discarded so a failed batch
                            // yields no partial rows; callers resume after the
                            // last row they actually received.
                            *guard = None;
                            return Err(e);
                        }
                        _ => state.receiver = None,
                    }
                }
                let n = size.min(state.ready.len());
                Ok(state.ready.drain(..n).collect())
            })
            .map_err(error)?;
        let result = PyList::empty(py);
        for row in rows {
            let mut values = Vec::with_capacity(row.len());
            for cell in row {
                let value: Py<PyAny> = match cell {
                    Cell::Empty => py.None(),
                    Cell::Int(v) => v.into_pyobject(py)?.into_any().unbind(),
                    Cell::Float(v) => v.into_pyobject(py)?.into_any().unbind(),
                    Cell::Bool(v) => PyBool::new(py, v).to_owned().into_any().unbind(),
                    Cell::Text(v) => PyString::new(py, &v).into_any().unbind(),
                    Cell::Formula(v) => PyString::new(py, &format!("={v}")).into_any().unbind(),
                    Cell::Error(v) => PyString::new(py, &v).into_any().unbind(),
                };
                values.push(value);
            }
            result.append(PyTuple::new(py, values)?)?;
        }
        Ok(result.unbind())
    }

    fn close(&self, py: Python<'_>) -> PyResult<()> {
        py.detach(|| {
            self.rows
                .lock()
                .map(|mut r| *r = None)
                .map_err(|_| error("reader lock poisoned"))
        })
    }
}

#[pyclass]
struct NativeWriter {
    workbook: Mutex<write::Workbook>,
}

#[pymethods]
impl NativeWriter {
    #[new]
    fn new() -> Self {
        Self {
            workbook: Mutex::new(write::Workbook::new()),
        }
    }

    fn create_sheet(&self, py: Python<'_>, name: String) -> PyResult<usize> {
        py.detach(|| {
            self.workbook
                .lock()
                .map_err(|_| "writer lock poisoned".into())
                .and_then(|mut w| w.create_sheet(&name))
        })
        .map_err(error)
    }

    fn append_rows(&self, py: Python<'_>, sheet: usize, rows: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut converted: Vec<Row> = Vec::new();
        for row in rows.try_iter()? {
            let mut values = Vec::new();
            let row = row?;
            if row.is_instance_of::<PyString>()
                || row.is_instance_of::<PyBytes>()
                || row.is_instance_of::<PyDict>()
            {
                return Err(PyTypeError::new_err(
                    "row must be an iterable of values, not a string, bytes, or mapping",
                ));
            }
            for value in row.try_iter()? {
                values.push(from_python(&value?)?);
                if values.len() > 16_384 {
                    return Err(error("column count exceeds Excel limits"));
                }
            }
            converted.push(values);
            if converted.len() > 65_536 {
                return Err(error("native append_rows batch exceeds 65536 rows"));
            }
        }
        py.detach(|| -> crate::Result<()> {
            let mut writer = self.workbook.lock().map_err(|_| "writer lock poisoned")?;
            for row in &converted {
                writer.append(sheet, row)?;
            }
            Ok(())
        })
        .map_err(error)
    }

    fn save(&self, py: Python<'_>, path: String, compression_level: u8) -> PyResult<()> {
        py.detach(|| {
            self.workbook
                .lock()
                .map_err(|_| "writer lock poisoned".into())
                .and_then(|mut w| w.save(path, compression_level))
        })
        .map_err(error)
    }
}

fn from_python(value: &Bound<'_, PyAny>) -> PyResult<Cell> {
    if value.is_none() {
        return Ok(Cell::Empty);
    }
    if value.is_instance_of::<PyBool>() {
        return Ok(Cell::Bool(value.extract()?));
    }
    if value.is_instance_of::<PyInt>() {
        return Ok(Cell::Int(value.extract()?));
    }
    if value.is_instance_of::<PyFloat>() {
        return Ok(Cell::Float(value.extract()?));
    }
    if let Ok(s) = value.cast::<PyString>() {
        let s = s.to_str()?;
        if s.starts_with('=') && s.len() > 1 {
            return Ok(Cell::Formula(s[1..].to_string()));
        }
        if [
            "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
        ]
        .contains(&s)
        {
            return Ok(Cell::Error(s.to_string()));
        }
        return Ok(Cell::Text(Arc::from(s)));
    }
    Err(PyTypeError::new_err(
        "supported values: None, bool, signed 64-bit int, float, str; dates and styled cells are not yet implemented",
    ))
}

#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<NativeWorkbook>()?;
    m.add_class::<NativeRows>()?;
    m.add_class::<NativeWriter>()?;
    Ok(())
}
