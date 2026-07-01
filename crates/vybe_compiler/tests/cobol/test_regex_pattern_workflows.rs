use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn inspect_tallying_all_occurrences_runtime() {
    let out = run_prints(&p(
        "01 WS-TEXT PIC X(20) VALUE \"ABCAABC\".\n01 WS-CNT PIC 9(2) VALUE 0.",
        "    INSPECT WS-TEXT TALLYING WS-CNT FOR ALL \"A\".\n    DISPLAY WS-CNT.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn inspect_replacing_first_occurrence_runtime() {
    let out = run_prints(&p(
        "01 WS-TEXT PIC X(10) VALUE \"BANANA\".",
        "    INSPECT WS-TEXT REPLACING FIRST \"A\" BY \"X\".\n    DISPLAY WS-TEXT.",
    ));
    assert_eq!(out, vec!["BXNANA"]);
}

#[test]
fn pattern_preprocessing_with_intrinsics_runtime() {
    let out = run_prints(&p(
        "01 WS-TEXT PIC X(20) VALUE \"  AbC123  \".\n01 WS-NORM PIC X(20).",
        "    MOVE FUNCTION TRIM(WS-TEXT) TO WS-NORM.\n    MOVE FUNCTION LOWER-CASE(WS-NORM) TO WS-NORM.\n    DISPLAY WS-NORM.",
    ));
    assert_eq!(out, vec!["abc123"]);
}

#[test]
fn unstring_tokenization_runtime() {
    let out = run_prints(&p(
        "01 SRC PIC X(20) VALUE \"AA-BB-CC\".\n01 A PIC X(4).\n01 B PIC X(4).\n01 C PIC X(4).",
        "    UNSTRING SRC DELIMITED BY \"-\" INTO A B C.\n    DISPLAY A.\n    DISPLAY B.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["AA", "BB", "CC"]);
}

#[test]
fn reference_modification_runtime() {
    let out = run_prints(&p(
        "01 WS-T PIC X(12) VALUE \"HELLOWORLD\".\n01 WS-S PIC X(5).",
        "    MOVE WS-T(6:5) TO WS-S.\n    DISPLAY WS-S.",
    ));
    assert_eq!(out, vec!["WORLD"]);
}
