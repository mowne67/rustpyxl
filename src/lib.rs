#[cfg(feature = "python")]
mod python;
pub mod read;
pub mod write;
mod xml;

pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error + Send + Sync>>;

#[derive(Debug, Clone, PartialEq)]
pub enum Cell {
    Empty,
    Int(i64),
    Float(f64),
    Bool(bool),
    Text(std::sync::Arc<str>),
    Formula(String),
    Error(String),
}

pub type Row = Vec<Cell>;
