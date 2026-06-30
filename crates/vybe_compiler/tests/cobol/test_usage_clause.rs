use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_usage_clauses() {
    compile_ok(&p(
        r#"
01 WS-A PIC 9(4) USAGE IS DISPLAY.
01 WS-B PIC 9(4) USAGE IS BINARY.
01 WS-C PIC 9(4) USAGE IS PACKED-DECIMAL.
01 WS-D PIC 9(4) USAGE IS COMPUTATIONAL.
01 WS-E PIC 9(4) USAGE IS COMP.
01 WS-F PIC 9(4) USAGE IS COMP-3.
"#,
        r#"
    ADD 5 TO WS-A WS-B WS-C WS-D WS-E WS-F.
"#,
    ));
}

#[test]
fn test_usage_floats() {
    compile_ok(&p(
        r#"
01 WS-FLOAT-S USAGE IS COMP-1.
01 WS-FLOAT-D USAGE IS COMP-2.
"#,
        r#"
    DISPLAY WS-FLOAT-S.
"#,
    ));
}

#[test]
fn test_usage_native_binary() {
    compile_ok(&p(
        r#"
01 WS-BIN PIC 9(4) USAGE IS COMP-5.
"#,
        r#"
    ADD 1 TO WS-BIN.
"#,
    ));
}

#[test]
fn test_usage_pointers() {
    compile_ok(&p(
        r#"
01 WS-PTR USAGE IS POINTER.
01 WS-FPTR USAGE IS FUNCTION-POINTER.
01 WS-PPTR USAGE IS PROCEDURE-POINTER.
"#,
        r#"
    SET WS-PTR TO NULL.
"#,
    ));
}

#[test]
fn test_usage_group_inheritance() {
    compile_ok(&p(
        r#"
01 WS-GROUP USAGE IS BINARY.
   05 WS-A PIC 9(4).
   05 WS-B PIC 9(4).
"#,
        r#"
    ADD 10 TO WS-A.
"#,
    ));
}
