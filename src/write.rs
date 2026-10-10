use crate::{Cell, Result, Row, xml};
use std::{
    fs::File,
    io::{BufWriter, Seek, SeekFrom, Write},
    path::Path,
};
use tempfile::NamedTempFile;
use zip::{CompressionMethod, ZipWriter, write::SimpleFileOptions};

const HEADER: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

struct Sheet {
    name: String,
    file: BufWriter<NamedTempFile>,
    rows: usize,
}

/// Rows are spooled to disk. No cell graph or shared-string table is retained.
pub struct Workbook {
    sheets: Vec<Sheet>,
    saved: bool,
    poisoned: bool,
    row_buf: String,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

impl Workbook {
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            saved: false,
            poisoned: false,
            row_buf: String::with_capacity(4096),
        }
    }

    pub fn create_sheet(&mut self, name: &str) -> Result<usize> {
        self.ensure_open()?;
        if name.is_empty()
            || name.chars().count() > 31
            || name.contains(['[', ']', ':', '*', '?', '/', '\\'])
            || !xml::valid_text(name)
        {
            return Err("sheet name must contain 1–31 characters without []:*?/\\ or XML control characters".into());
        }
        if self
            .sheets
            .iter()
            .any(|s| s.name.to_lowercase() == name.to_lowercase())
        {
            return Err("duplicate sheet name".into());
        }
        let mut file = BufWriter::with_capacity(64 * 1024, NamedTempFile::new()?);
        write!(file, "{HEADER}<worksheet xmlns=\"{NS}\"><sheetData>")?;
        self.sheets.push(Sheet {
            name: name.into(),
            file,
            rows: 0,
        });
        Ok(self.sheets.len() - 1)
    }

    fn ensure_open(&self) -> Result<()> {
        if self.poisoned {
            return Err("writer cannot continue after an I/O failure".into());
        }
        if self.saved {
            return Err("write-only workbook has already been saved".into());
        }
        Ok(())
    }

    pub fn append(&mut self, sheet: usize, row: &Row) -> Result<()> {
        use std::fmt::Write as _;
        self.ensure_open()?;
        let sheet = self
            .sheets
            .get_mut(sheet)
            .ok_or("sheet index out of range")?;
        if sheet.rows >= 1_048_576 || row.len() > 16_384 {
            return Err("row or column count exceeds Excel limits".into());
        }
        let number = sheet.rows + 1;
        self.row_buf.clear();
        write!(self.row_buf, "<row r=\"{number}\">")?;
        for (index, cell) in row.iter().enumerate() {
            if matches!(cell, Cell::Empty) {
                continue;
            }
            let coord = xml::Coordinate(index + 1, number);
            match cell {
                Cell::Int(v) => write!(self.row_buf, "<c r=\"{coord}\"><v>{v}</v></c>")?,
                Cell::Float(v) => {
                    if !v.is_finite() {
                        return Err("NaN and infinity cannot be written to XLSX".into());
                    }
                    write!(self.row_buf, "<c r=\"{coord}\"><v>{v}</v></c>")?;
                }
                Cell::Bool(v) => write!(
                    self.row_buf,
                    "<c r=\"{coord}\" t=\"b\"><v>{}</v></c>",
                    u8::from(*v)
                )?,
                Cell::Text(v) => {
                    if !xml::valid_text(v) {
                        return Err("string contains illegal XML characters".into());
                    }
                    if v.chars().count() > 32_767 {
                        return Err("string exceeds Excel's 32767 character limit".into());
                    }
                    if v.is_empty() {
                        write!(self.row_buf, "<c r=\"{coord}\" t=\"inlineStr\"></c>")?;
                    } else {
                        let escaped = quick_xml::escape::escape(v.as_ref());
                        write!(
                            self.row_buf,
                            "<c r=\"{coord}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{}</t></is></c>",
                            if escaped.contains('\r') {
                                escaped.replace('\r', "&#13;").into()
                            } else {
                                escaped
                            }
                        )?;
                    }
                }
                Cell::Formula(v) => {
                    if !xml::valid_text(v) {
                        return Err("formula contains illegal XML characters".into());
                    }
                    if v.chars().count() >= 32_767 {
                        return Err(
                            "formula including '=' exceeds Excel's 32767 character limit".into(),
                        );
                    }
                    write!(
                        self.row_buf,
                        "<c r=\"{coord}\"><f>{}</f><v></v></c>",
                        quick_xml::escape::escape(v)
                    )?;
                }
                Cell::Error(v) => {
                    if ![
                        "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
                    ]
                    .contains(&v.as_str())
                    {
                        return Err("invalid Excel error token".into());
                    }
                    write!(self.row_buf, "<c r=\"{coord}\" t=\"e\"><v>{v}</v></c>")?;
                }
                Cell::Empty => (),
            }
        }
        self.row_buf.push_str("</row>");
        if let Err(e) = sheet.file.write_all(self.row_buf.as_bytes()) {
            self.poisoned = true;
            return Err(e.into());
        }
        sheet.rows += 1;
        Ok(())
    }

    /// Serialize into a sibling temporary file; publish only a completed archive.
    pub fn save(&mut self, path: impl AsRef<Path>, compression_level: u8) -> Result<()> {
        self.ensure_open()?;
        if self.sheets.is_empty() {
            return Err("a workbook needs at least one sheet".into());
        }
        if compression_level > 9 {
            return Err("compression_level must be in 0..9".into());
        }
        let path = path.as_ref();
        let parent = path
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let mut output = NamedTempFile::new_in(parent)?;
        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .compression_level(Some(compression_level as i64));
        let mut zip = ZipWriter::new(output.as_file_mut());
        let mut types = format!(
            "{HEADER}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        );
        for i in 1..=self.sheets.len() {
            types.push_str(&format!("<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"));
        }
        types.push_str("</Types>");
        entry(&mut zip, "[Content_Types].xml", &types, options)?;
        entry(
            &mut zip,
            "_rels/.rels",
            &format!(
                "{HEADER}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{REL}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
            ),
            options,
        )?;
        let mut workbook = format!("{HEADER}<workbook xmlns=\"{NS}\" xmlns:r=\"{REL}\"><sheets>");
        let mut rels = format!(
            "{HEADER}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        );
        for (i, sheet) in self.sheets.iter().enumerate() {
            let id = i + 1;
            workbook.push_str(&format!(
                "<sheet name=\"{}\" sheetId=\"{id}\" r:id=\"rId{id}\"/>",
                quick_xml::escape::escape(&sheet.name)
                    .replace('\t', "&#9;")
                    .replace('\n', "&#10;")
            ));
            rels.push_str(&format!("<Relationship Id=\"rId{id}\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet{id}.xml\"/>"));
        }
        workbook.push_str("</sheets><calcPr fullCalcOnLoad=\"1\"/></workbook>");
        rels.push_str("</Relationships>");
        entry(&mut zip, "xl/workbook.xml", &workbook, options)?;
        entry(&mut zip, "xl/_rels/workbook.xml.rels", &rels, options)?;
        // Do not modify spool contents: a failed save can be retried safely.
        for (i, sheet) in self.sheets.iter_mut().enumerate() {
            sheet.file.flush()?;
            let mut source = File::open(sheet.file.get_ref().path())?;
            source.seek(SeekFrom::Start(0))?;
            zip.start_file(format!("xl/worksheets/sheet{}.xml", i + 1), options)?;
            std::io::copy(&mut source, &mut zip)?;
            zip.write_all(b"</sheetData></worksheet>")?;
        }
        zip.finish()?;
        output.as_file().sync_all()?;
        output.persist(path)?;
        self.saved = true;
        self.sheets.clear();
        Ok(())
    }
}

fn entry<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    path: &str,
    xml: &str,
    options: SimpleFileOptions,
) -> Result<()> {
    zip.start_file(path, options)?;
    zip.write_all(xml.as_bytes())?;
    Ok(())
}
