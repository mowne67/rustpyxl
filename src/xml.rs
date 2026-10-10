use crate::Result;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event, attributes::Attribute};
use std::borrow::Cow;
use std::io::BufRead;

pub fn attr(e: &BytesStart<'_>, key: &str) -> Result<Option<String>> {
    Ok(attr_ref(e, key)?.map(Cow::into_owned))
}

pub fn attr_ref<'a>(e: &'a BytesStart<'_>, key: &str) -> Result<Option<Cow<'a, str>>> {
    for a in e.attributes() {
        let a = a?;
        if a.key.as_ref() == key {
            return Ok(Some(value(a)?));
        }
    }
    Ok(None)
}

pub fn value(a: Attribute<'_>) -> Result<Cow<'_, str>> {
    Ok(match a.value {
        Cow::Borrowed(v) => quick_xml::escape::unescape(v)?,
        Cow::Owned(v) => Cow::Owned(quick_xml::escape::unescape(&v)?.into_owned()),
    })
}

pub fn text<'a>(e: &'a Event<'_>) -> Result<Option<Cow<'a, str>>> {
    Ok(match e {
        Event::Text(t) => Some(t.xml_content(quick_xml::XmlVersion::Implicit1_0)),
        Event::CData(t) => Some(t.xml_content(quick_xml::XmlVersion::Implicit1_0)),
        Event::GeneralRef(r) => {
            let name = r.as_ref();
            Some(Cow::Owned(
                quick_xml::escape::unescape(&format!("&{name};"))?.into_owned(),
            ))
        }
        _ => None,
    })
}

pub fn reader<R: BufRead>(input: R) -> Reader<R> {
    let mut r = Reader::from_reader(input);
    r.config_mut().expand_empty_elements = true;
    r
}

pub fn column(reference: &str) -> Result<usize> {
    let mut col = 0usize;
    let mut i = 0;
    for b in reference.bytes() {
        if !b.is_ascii_uppercase() {
            break;
        }
        col = col * 26 + (b - b'A' + 1) as usize;
        i += 1;
        if col > 16_384 {
            return Err("column exceeds Excel's XFD limit".into());
        }
    }
    let row = match reference[i..].parse::<usize>() {
        Ok(row) if col > 0 => row,
        _ => return Err(format!("invalid cell reference: {reference}").into()),
    };
    if !(1..=1_048_576).contains(&row) {
        return Err("row exceeds Excel's limit".into());
    }
    Ok(col)
}

/// Cell reference such as `B7`, formatted without a heap allocation.
pub struct Coordinate(pub usize, pub usize);

impl std::fmt::Display for Coordinate {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let (mut col, mut letters, mut i) = (self.0, [0u8; 3], 3);
        while col > 0 {
            col -= 1;
            i -= 1;
            letters[i] = b'A' + (col % 26) as u8;
            col /= 26;
        }
        f.write_str(std::str::from_utf8(&letters[i..]).unwrap())?;
        write!(f, "{}", self.1)
    }
}

pub fn valid_text(s: &str) -> bool {
    s.chars()
        .all(|c| matches!(c, '\t' | '\n' | '\r') || c >= '\u{20}')
        && !s.contains(['\u{fffe}', '\u{ffff}'])
}
