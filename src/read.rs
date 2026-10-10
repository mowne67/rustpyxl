use crate::{Cell, Result, Row, xml};
use quick_xml::events::Event;
use std::{
    collections::HashMap,
    fs::File,
    io::{self, BufReader, Read, Seek, SeekFrom},
    path::{Path, PathBuf},
    sync::Arc,
};
use zip::{CompressionMethod, ZipArchive};

/// Owns only the compressed member stream, not the entire ZIP or worksheet.
struct Member {
    inner: Box<dyn Read + Send>,
    crc: crc32fast::Hasher,
    expected_crc: u32,
    size: u64,
    count: u64,
    checked: bool,
}

impl Read for Member {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        if buf.is_empty() {
            return Ok(0);
        }
        let n = self.inner.read(buf)?;
        self.count += n as u64;
        if self.count > self.size {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "ZIP member exceeds declared size",
            ));
        }
        self.crc.update(&buf[..n]);
        if n == 0 && !self.checked {
            self.checked = true;
            if self.count != self.size || self.crc.clone().finalize() != self.expected_crc {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "ZIP member size or CRC mismatch",
                ));
            }
        }
        Ok(n)
    }
}

fn member(path: &Path, name: &str) -> Result<BufReader<Member>> {
    let mut archive = ZipArchive::new(File::open(path)?)?;
    let entry = archive.by_name(name)?;
    if entry.encrypted() {
        return Err("encrypted XLSX files are unsupported".into());
    }
    let mut file = File::open(path)?;
    file.seek(SeekFrom::Start(
        entry.data_start().ok_or("ZIP member lacks data offset")?,
    ))?;
    let compressed = file.take(entry.compressed_size());
    let inner: Box<dyn Read + Send> = match entry.compression() {
        CompressionMethod::Stored => Box::new(compressed),
        CompressionMethod::Deflated => Box::new(flate2::read::DeflateDecoder::new(compressed)),
        _ => return Err("unsupported ZIP compression method".into()),
    };
    Ok(BufReader::with_capacity(
        64 * 1024,
        Member {
            inner,
            crc: crc32fast::Hasher::new(),
            expected_crc: entry.crc32(),
            size: entry.size(),
            count: 0,
            checked: false,
        },
    ))
}

#[derive(Debug, Clone)]
pub struct Sheet {
    pub name: String,
    pub path: String,
}

pub struct Workbook {
    path: PathBuf,
    pub sheets: Vec<Sheet>,
    pub shared_strings_path: Option<String>,
    shared: Arc<Vec<Arc<str>>>,
    unsupported_styles: Arc<Vec<bool>>,
}

impl Workbook {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref().to_path_buf();
        let mut rels = HashMap::new();
        let mut shared_path = None;
        let mut styles_path = None;
        let mut r = xml::reader(member(&path, "xl/_rels/workbook.xml.rels")?);
        let mut buf = Vec::new();
        loop {
            match r.read_event_into(&mut buf)? {
                Event::Start(e) if e.local_name().as_ref() == "Relationship" => {
                    let id = xml::attr(&e, "Id")?.ok_or("relationship missing Id")?;
                    let target = xml::attr(&e, "Target")?.ok_or("relationship missing Target")?;
                    let mode = xml::attr(&e, "TargetMode")?;
                    if mode.as_deref() == Some("External") {
                        continue;
                    }
                    let full = normalize_target(&target)?;
                    let kind = xml::attr(&e, "Type")?.unwrap_or_default();
                    if kind.ends_with("/sharedStrings") {
                        shared_path = Some(full.clone());
                    }
                    if kind.ends_with("/styles") {
                        styles_path = Some(full.clone());
                    }
                    rels.insert(id, full);
                }
                Event::Eof => break,
                _ => (),
            }
            buf.clear();
        }
        let mut sheets = Vec::new();
        let mut r = xml::reader(member(&path, "xl/workbook.xml")?);
        loop {
            match r.read_event_into(&mut buf)? {
                Event::Start(e) if e.local_name().as_ref() == "sheet" => {
                    let name = xml::attr(&e, "name")?.ok_or("sheet missing name")?;
                    let mut id = None;
                    for a in e.attributes() {
                        let a = a?;
                        if a.key.local_name().as_ref() == "id" {
                            id = Some(a.value.to_string());
                        }
                    }
                    let sheet_path = rels
                        .get(&id.ok_or("sheet missing relationship")?)
                        .ok_or("missing sheet relationship")?
                        .clone();
                    sheets.push(Sheet {
                        name,
                        path: sheet_path,
                    });
                }
                Event::Eof => break,
                _ => (),
            }
            buf.clear();
        }
        let shared = match &shared_path {
            Some(p) => shared_strings(member(&path, p)?)?,
            None => Vec::new(),
        };
        let styles = match styles_path {
            Some(p) => unsupported_styles(member(&path, &p)?)?,
            None => Vec::new(),
        };
        Ok(Self {
            path,
            sheets,
            shared_strings_path: shared_path,
            shared: Arc::new(shared),
            unsupported_styles: Arc::new(styles),
        })
    }

    pub fn rows(&self, sheet: usize, data_only: bool) -> Result<Rows> {
        let sheet = self.sheets.get(sheet).ok_or("sheet index out of range")?;
        Ok(Rows {
            reader: xml::reader(member(&self.path, &sheet.path)?),
            buf: Vec::with_capacity(4096),
            shared: self.shared.clone(),
            unsupported_styles: self.unsupported_styles.clone(),
            data_only,
            width: 0,
            last_row: 0,
            end_row: 0,
            pending: None,
            eof: false,
        })
    }

    pub fn shared_len(&self) -> usize {
        self.shared.len()
    }

    pub fn shared_string(&self, index: usize) -> Option<&str> {
        self.shared.get(index).map(|s| s.as_ref())
    }
}

fn normalize_target(target: &str) -> Result<String> {
    let base = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        format!("xl/{target}")
    };
    let mut parts = Vec::new();
    for p in base.split('/') {
        match p {
            "" | "." => (),
            ".." => {
                parts.pop().ok_or("relationship escapes package root")?;
            }
            p => parts.push(p),
        }
    }
    Ok(parts.join("/"))
}

fn shared_strings(input: BufReader<Member>) -> Result<Vec<Arc<str>>> {
    let mut r = xml::reader(input);
    let mut buf = Vec::new();
    let mut strings = Vec::new();
    let mut value = String::new();
    let mut in_text = false;
    let mut phonetic = false;
    loop {
        let e = r.read_event_into(&mut buf)?;
        match &e {
            Event::Start(e) if e.local_name().as_ref() == "si" => value.clear(),
            Event::Start(e) if e.local_name().as_ref() == "rPh" => phonetic = true,
            Event::End(e) if e.local_name().as_ref() == "rPh" => phonetic = false,
            Event::Start(e) if e.local_name().as_ref() == "t" => in_text = !phonetic,
            Event::End(e) if e.local_name().as_ref() == "t" => in_text = false,
            Event::End(e) if e.local_name().as_ref() == "si" => {
                if value.contains("x005F_") {
                    strings.push(Arc::from(value.replace("x005F_", "")));
                } else {
                    strings.push(Arc::from(value.as_str()));
                }
            }
            Event::Eof => break,
            _ if in_text => {
                if let Some(s) = xml::text(&e)? {
                    value.push_str(&s);
                }
            }
            _ => (),
        }
        buf.clear();
    }
    Ok(strings)
}

// Numeric dates and custom formats need a style engine. Fail explicitly rather
// than silently returning the wrong Python type in this initial milestone.
fn unsupported_styles(input: BufReader<Member>) -> Result<Vec<bool>> {
    let mut r = xml::reader(input);
    let mut buf = Vec::new();
    let mut xfs = false;
    let mut result = Vec::new();
    loop {
        match r.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == "cellXfs" => xfs = true,
            Event::End(e) if e.local_name().as_ref() == "cellXfs" => xfs = false,
            Event::Start(e) if xfs && e.local_name().as_ref() == "xf" => {
                let id: u32 = xml::attr(&e, "numFmtId")?
                    .unwrap_or_else(|| "0".into())
                    .parse()?;
                result.push(matches!(id, 14..=22 | 27..=36 | 45..=47 | 50..=58 | 164..));
            }
            Event::Eof => break,
            _ => (),
        }
        buf.clear();
    }
    Ok(result)
}

pub struct Rows {
    reader: quick_xml::Reader<BufReader<Member>>,
    buf: Vec<u8>,
    shared: Arc<Vec<Arc<str>>>,
    unsupported_styles: Arc<Vec<bool>>,
    data_only: bool,
    width: usize,
    last_row: usize,
    end_row: usize,
    pending: Option<(usize, Row)>,
    eof: bool,
}

impl Rows {
    pub fn next_row(&mut self) -> Result<Option<Row>> {
        if let Some((number, _)) = &self.pending {
            if self.last_row + 1 < *number {
                self.last_row += 1;
                return Ok(Some(vec![Cell::Empty; self.width]));
            }
            let (number, row) = self.pending.take().unwrap();
            self.last_row = number;
            return Ok(Some(row));
        }
        if self.eof {
            return Ok(None);
        }
        let mut row = Vec::with_capacity(self.width);
        let mut number = self.last_row + 1;
        let mut col = 0;
        let mut kind = String::new();
        let mut value = String::new();
        let mut formula = String::new();
        let mut field = 0;
        let mut has_formula = false;
        let mut styled = false;
        let mut in_cell = false;
        let mut in_row = false;
        let mut phonetic = false;
        let mut has_inline_text = false;
        loop {
            let e = self.reader.read_event_into(&mut self.buf)?;
            match &e {
                Event::Start(e) => match e.local_name().as_ref() {
                    "dimension" => {
                        if let Some(reference) = xml::attr(e, "ref")? {
                            let end = reference.rsplit(':').next().unwrap();
                            self.width = xml::column(end)?;
                            self.end_row = end
                                .trim_start_matches(|c: char| c.is_ascii_uppercase())
                                .parse()?;
                        }
                    }
                    "row" => {
                        in_row = true;
                        number = xml::attr_ref(e, "r")?
                            .map(|s| s.parse())
                            .transpose()?
                            .unwrap_or(number);
                        if number <= self.last_row || number > 1_048_576 {
                            return Err("invalid or non-monotonic row index".into());
                        }
                    }
                    "c" if in_row => {
                        in_cell = true;
                        // One attribute pass per cell; this is the hottest loop.
                        let mut reference = None;
                        let mut style = 0usize;
                        kind.clear();
                        // Skipping duplicate-attribute checks avoids an allocation per cell.
                        for a in e.attributes().with_checks(false) {
                            let a = a?;
                            match a.key.as_ref() {
                                "r" => reference = Some(xml::column(&xml::value(a)?)?),
                                "t" => kind.push_str(&xml::value(a)?),
                                "s" => style = xml::value(a)?.parse()?,
                                _ => (),
                            }
                        }
                        col = reference.unwrap_or(col + 1);
                        if col == 0 || col > 16_384 || col <= row.len() {
                            return Err("invalid or non-monotonic cell column".into());
                        }
                        if style > 0 && style >= self.unsupported_styles.len() {
                            return Err("cell style index out of range".into());
                        }
                        styled = self.unsupported_styles.get(style).copied().unwrap_or(false);
                        value.clear();
                        formula.clear();
                        has_formula = false;
                        has_inline_text = false;
                        field = 0;
                    }
                    "f" if in_cell => {
                        if let Some(t) = xml::attr_ref(e, "t")?
                            && t != "normal"
                            && !self.data_only
                        {
                            return Err(format!(
                                "unsupported {t} formula; use data_only=True for cached results"
                            )
                            .into());
                        }
                        has_formula = true;
                        field = 2;
                    }
                    "rPh" if in_cell => phonetic = true,
                    "v" | "t" if in_cell && !phonetic => {
                        if e.local_name().as_ref() == "t" {
                            has_inline_text = true;
                        }
                        field = 1;
                    }
                    _ => (),
                },
                Event::End(e) => match e.local_name().as_ref() {
                    "v" | "t" | "f" => field = 0,
                    "rPh" => phonetic = false,
                    "c" if in_cell => {
                        let cell = if kind == "inlineStr" && !has_inline_text {
                            Cell::Empty
                        } else if has_formula && !self.data_only {
                            Cell::Formula(formula.clone())
                        } else {
                            decode_cell(&kind, &value, &self.shared, styled)?
                        };
                        row.resize(col - 1, Cell::Empty);
                        row.push(cell);
                        in_cell = false;
                    }
                    "row" if in_row => {
                        row.resize(row.len().max(self.width), Cell::Empty);
                        self.pending = Some((number, row));
                        self.buf.clear();
                        return self.next_row();
                    }
                    _ => (),
                },
                Event::Eof => {
                    if in_row || in_cell {
                        return Err("truncated worksheet XML".into());
                    }
                    self.eof = true;
                    self.buf.clear();
                    return self.next_row();
                }
                _ if field != 0 => {
                    if let Some(s) = xml::text(&e)? {
                        if field == 1 {
                            value.push_str(&s);
                        } else {
                            formula.push_str(&s);
                        }
                    }
                }
                _ => (),
            }
            self.buf.clear();
        }
    }

    pub fn batch(&mut self, size: usize) -> Result<Vec<Row>> {
        let mut rows = Vec::with_capacity(size);
        for _ in 0..size {
            match self.next_row()? {
                Some(r) => rows.push(r),
                None => break,
            }
        }
        Ok(rows)
    }
}

fn decode_cell(kind: &str, value: &str, shared: &[Arc<str>], styled: bool) -> Result<Cell> {
    if kind == "str" && value.is_empty() {
        return Ok(Cell::Empty);
    }
    if matches!(kind, "inlineStr" | "str") {
        return Ok(Cell::Text(Arc::from(value)));
    }
    if value.is_empty() {
        return Ok(Cell::Empty);
    }
    Ok(match kind {
        "s" => Cell::Text(
            shared
                .get(value.parse::<usize>()?)
                .ok_or("shared string index out of range")?
                .clone(),
        ),
        "b" => match value {
            "0" => Cell::Bool(false),
            "1" => Cell::Bool(true),
            _ => return Err("invalid boolean cell".into()),
        },
        "e" => Cell::Error(value.to_string()),
        "d" => return Err("ISO date cells are not supported in this milestone".into()),
        "" | "n" => {
            if styled {
                return Err(
                    "numeric date or custom number format requires the future style engine".into(),
                );
            }
            match value.parse::<i64>() {
                Ok(v) => Cell::Int(v),
                Err(_) => {
                    if !value.contains(['.', 'e', 'E']) {
                        return Err(
                            "invalid integer cell or value exceeds signed 64-bit range".into()
                        );
                    }
                    let v = value.parse::<f64>()?;
                    if !v.is_finite() {
                        return Err("non-finite numeric cell".into());
                    }
                    Cell::Float(v)
                }
            }
        }
        _ => return Err(format!("unsupported cell type: {kind}").into()),
    })
}
