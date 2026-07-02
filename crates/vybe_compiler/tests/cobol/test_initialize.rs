use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_initialize_elementary() {
    let output = run_prints(&p(
        r#"
01 WS-NUM PIC 9(3) VALUE 123.
01 WS-STR PIC X(5) VALUE "HELLO".
"#,
        r#"
    INITIALIZE WS-NUM WS-STR.
    DISPLAY WS-NUM.
    DISPLAY WS-STR.
"#,
    ));
    assert_eq!(output, vec!["000", "     "]);
}

#[test]
fn test_initialize_group() {
    let output = run_prints(&p(
        r#"
01 WS-GROUP.
   05 WS-A PIC 9(3) VALUE 123.
   05 WS-B PIC X(5) VALUE "HELLO".
"#,
        r#"
    INITIALIZE WS-GROUP.
    DISPLAY WS-A.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["000", "     "]);
}

#[test]
fn test_initialize_replacing() {
    let output = run_prints(&p(
        r#"
01 WS-GROUP.
   05 WS-A PIC 9(3) VALUE 123.
   05 WS-B PIC X(5) VALUE "HELLO".
"#,
        r#"
    INITIALIZE WS-GROUP REPLACING NUMERIC BY 5.
    DISPLAY WS-A.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["005", "HELLO"]);
}

#[test]
fn test_initialize_replacing_alpha() {
    let output = run_prints(&p(
        r#"
01 WS-GROUP.
   05 WS-A PIC 9(3) VALUE 123.
   05 WS-B PIC X(5) VALUE "HELLO".
"#,
        r#"
    INITIALIZE WS-GROUP REPLACING ALPHANUMERIC BY "WORLD".
    DISPLAY WS-A.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["123", "WORLD"]);
}

#[test]
fn test_initialize_table() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 3 TIMES VALUE 100.
"#,
        r#"
    INITIALIZE WS-TABLE.
    DISPLAY WS-ITEM(1).
    DISPLAY WS-ITEM(2).
"#,
    ));
    assert_eq!(output, vec!["000", "000"]);
}
