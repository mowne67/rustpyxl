use _native::{Cell, read, write};

#[test]
fn rust_core_roundtrip() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("core.xlsx");
    let mut workbook = write::Workbook::new();
    let sheet = workbook.create_sheet("Rust").unwrap();
    let row = vec![
        Cell::Int(7),
        Cell::Empty,
        Cell::Text(" <&> 🦀 ".into()),
        Cell::Bool(true),
        Cell::Formula("A1+1".into()),
    ];
    workbook.append(sheet, &row).unwrap();
    workbook.save(&path, 6).unwrap();
    let source = read::Workbook::open(&path).unwrap();
    let mut rows = source.rows(0, false).unwrap();
    assert_eq!(rows.next_row().unwrap(), Some(row));
    assert_eq!(rows.next_row().unwrap(), None);
    let mut values = source.rows(0, true).unwrap();
    assert_eq!(values.next_row().unwrap().unwrap()[4], Cell::Empty);
}

#[test]
fn rejected_row_does_not_partially_append() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("rejected.xlsx");
    let mut workbook = write::Workbook::new();
    let sheet = workbook.create_sheet("Sheet").unwrap();
    assert!(
        workbook
            .append(sheet, &vec![Cell::Int(1), Cell::Float(f64::NAN)])
            .is_err()
    );
    workbook.append(sheet, &vec![Cell::Int(2)]).unwrap();
    workbook.save(&path, 6).unwrap();
    let source = read::Workbook::open(&path).unwrap();
    assert_eq!(
        source.rows(0, false).unwrap().batch(10).unwrap(),
        vec![vec![Cell::Int(2)]]
    );
}
