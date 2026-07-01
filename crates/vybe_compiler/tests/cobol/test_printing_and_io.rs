use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn display_literal_and_variable_compiles() {
    compile_ok(&p(
        "01 WS-NAME PIC X(10) VALUE \"ALICE\".",
        "    DISPLAY \"Hello\".\n    DISPLAY WS-NAME.",
    ));
}

#[test]
fn display_multiple_items_compiles() {
    compile_ok(&p(
        "01 WS-NAME PIC X(10) VALUE \"BOB\".\n01 WS-AGE PIC 9(2) VALUE 42.",
        "    DISPLAY \"Name: \" WS-NAME \" Age: \" WS-AGE.",
    ));
}

#[test]
fn accept_text_and_date_compiles() {
    compile_ok(&p(
        "01 WS-NAME PIC X(20).\n01 WS-DATE PIC X(8).",
        "    ACCEPT WS-NAME.\n    ACCEPT WS-DATE FROM DATE.",
    ));
}

#[test]
fn open_close_and_read_write_compiles() {
    compile_ok(&p(
        "01 WS-REC PIC X(80).",
        "    OPEN INPUT WS-FILE.\n    READ WS-FILE INTO WS-REC.\n    WRITE WS-REC FROM WS-REC.\n    CLOSE WS-FILE.",
    ));
}

#[test]
fn display_multiple_items_runtime_formats_output() {
    let output = run_prints(&p(
        "01 WS-NAME PIC X(5) VALUE \"ALICE\".\n01 WS-AGE PIC 9(2) VALUE 31.",
        "    DISPLAY \"Name=\" WS-NAME \" Age=\" WS-AGE.",
    ));
    assert_eq!(output, vec!["Name=ALICE Age=31"]);
}

#[test]
fn display_literal_runtime_prints_exact_text() {
    let output = run_prints(&p("", "    DISPLAY \"COBOL-IO\"."));
    assert_eq!(output, vec!["COBOL-IO"]);
}

#[test]
fn display_sequence_runtime_preserves_order() {
    let output = run_prints(&p(
        "01 A PIC X(3) VALUE \"ONE\".\n01 B PIC X(3) VALUE \"TWO\".",
        "    DISPLAY A.\n    DISPLAY B.",
    ));
    assert_eq!(output, vec!["ONE", "TWO"]);
}

#[test]
fn accept_from_date_then_display_compiles() {
    compile_ok(&p(
        "01 WS-DATE PIC X(8).",
        "    ACCEPT WS-DATE FROM DATE YYYYMMDD.\n    DISPLAY WS-DATE.",
    ));
}
