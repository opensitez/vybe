use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn numeric_item_with_value_compiles() {
    let out = run_prints(&p(
        "01 WS-NUM PIC 9(3) VALUE 5.",
        "    MOVE 7 TO WS-NUM.\n    DISPLAY WS-NUM.",
    ));
    assert_eq!(out, vec!["7"]);
}

#[test]
fn signed_numeric_item_compiles() {
    let out = run_prints(&p(
        "01 WS-NUM PIC S9(3) VALUE -3.",
        "    MOVE -4 TO WS-NUM.\n    DISPLAY WS-NUM.",
    ));
    assert_eq!(out, vec!["-4"]);
}

#[test]
fn edited_numeric_item_compiles() {
    let out = run_prints(&p(
        "01 WS-ED PIC ZZ9.99 VALUE 12.34.",
        "    MOVE 7.5 TO WS-ED.\n    DISPLAY WS-ED.",
    ));
    assert_eq!(out, vec!["7.5"]);
}

#[test]
fn alphanumeric_item_compiles() {
    let out = run_prints(&p(
        "01 WS-TXT PIC X(12) VALUE \"HELLO\".",
        "    MOVE \"WORLD\" TO WS-TXT.\n    DISPLAY WS-TXT.",
    ));
    assert_eq!(out, vec!["WORLD"]);
}

#[test]
fn alphanumeric_editted_item_compiles() {
    let out = run_prints(&p(
        "01 WS-TXT PIC X(12) VALUE \"ABC\".",
        "    MOVE \"123\" TO WS-TXT.\n    DISPLAY WS-TXT.",
    ));
    assert_eq!(out, vec!["123"]);
}

#[test]
fn group_item_compiles() {
    let out = run_prints(&p(
        "01 WS-REC.\n   05 WS-A PIC X(3).\n   05 WS-B PIC 9(2).",
        "    MOVE \"XYZ\" TO WS-A.\n    MOVE 99 TO WS-B.\n    DISPLAY WS-A.\n    DISPLAY WS-B.",
    ));
    assert_eq!(out, vec!["XYZ", "99"]);
}

#[test]
fn nested_group_item_compiles() {
    let out = run_prints(&p(
        "01 WS-REC.\n   05 WS-INFO.\n      10 WS-NAME PIC X(5).\n      10 WS-AGE PIC 9(2).",
        "    MOVE \"BOB\" TO WS-NAME.\n    MOVE 42 TO WS-AGE.\n    DISPLAY WS-NAME.\n    DISPLAY WS-AGE.",
    ));
    assert_eq!(out, vec!["BOB", "42"]);
}

#[test]
fn occurs_clause_compiles() {
    let out = run_prints(&p(
        "01 WS-TABLE PIC X(3) OCCURS 5 TIMES.",
        "    MOVE \"A\" TO WS-TABLE(1).\n    MOVE \"B\" TO WS-TABLE(2).\n    MOVE \"C\" TO WS-TABLE(3).\n    DISPLAY WS-TABLE(1).\n    DISPLAY WS-TABLE(2).\n    DISPLAY WS-TABLE(3).",
    ));
    assert_eq!(out, vec!["A", "B", "C"]);
}

#[test]
fn occurs_with_index_compiles() {
    let out = run_prints(&p(
        "01 WS-TABLE PIC 9(2) OCCURS 4 TIMES.\n01 WS-IDX PIC 9(1) VALUE 1.",
        "    MOVE 11 TO WS-TABLE(1).\n    MOVE 22 TO WS-TABLE(2).\n    MOVE 33 TO WS-TABLE(3).\n    MOVE 44 TO WS-TABLE(4).\n    MOVE 2 TO WS-IDX.\n    DISPLAY WS-TABLE(WS-IDX).",
    ));
    assert_eq!(out, vec!["22"]);
}

#[test]
fn redefines_item_compiles() {
    let out = run_prints(&p(
        "01 WS-BUF PIC X(10).\n01 WS-NUM REDEFINES WS-BUF PIC 9(10).",
        "    MOVE 42 TO WS-NUM.\n    MOVE WS-NUM TO WS-BUF.\n    DISPLAY WS-BUF.",
    ));
    assert_eq!(out, vec!["0000000042"]);
}

#[test]
fn redefines_group_item_compiles() {
    let out = run_prints(&p(
        "01 WS-BUF PIC X(20).\n01 WS-GRP REDEFINES WS-BUF.\n   05 WS-A PIC X(4).\n   05 WS-B PIC 9(2).",
        "    MOVE \"HI\" TO WS-A.\n    MOVE 7 TO WS-B.\n    DISPLAY WS-A.\n    DISPLAY WS-B.",
    ));
    assert_eq!(out, vec!["HI", "7"]);
}

#[test]
fn value_clause_string_compiles() {
    let out = run_prints(&p("01 WS-TXT PIC X(5) VALUE \"A\".", "    DISPLAY WS-TXT."));
    assert_eq!(out, vec!["A"]);
}

#[test]
fn value_clause_numeric_compiles() {
    let out = run_prints(&p("01 WS-NUM PIC 9(3) VALUE 100.", "    DISPLAY WS-NUM."));
    assert_eq!(out, vec!["100"]);
}

#[test]
fn value_clause_spaces_compiles() {
    let out = run_prints(&p(
        "01 WS-TXT PIC X(5) VALUE SPACES.",
        "    MOVE \"X\" TO WS-TXT.\n    DISPLAY WS-TXT.",
    ));
    assert_eq!(out, vec!["X"]);
}

#[test]
fn value_clause_zeros_compiles() {
    let out = run_prints(&p(
        "01 WS-NUM PIC 9(5) VALUE ZEROS.",
        "    MOVE 12345 TO WS-NUM.\n    DISPLAY WS-NUM.",
    ));
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn blank_when_zero_clause_compiles() {
    let out = run_prints(&p(
        "01 WS-NUM PIC 9(3) BLANK WHEN ZERO VALUE 0.\n01 WS-TMP PIC X(3).\n01 WS-OUT PIC X.",
        "    MOVE WS-NUM TO WS-TMP.\n    IF WS-TMP = \"   \"\n        MOVE 'Y' TO WS-OUT\n    ELSE\n        MOVE 'N' TO WS-OUT\n    END-IF.\n    DISPLAY WS-OUT.",
    ));
    assert_eq!(out, vec!["Y"]);
}

#[test]
fn just_justified_right_clause_compiles() {
    let out = run_prints(&p(
        "01 WS-TXT PIC X(5) JUSTIFIED RIGHT VALUE \"A\".",
        "    MOVE \"B\" TO WS-TXT.\n    DISPLAY WS-TXT.",
    ));
    assert_eq!(out, vec!["B"]);
}

#[test]
fn usage_display_item_compiles() {
    let out = run_prints(&p(
        "01 WS-NUM PIC 9(3) USAGE DISPLAY VALUE 5.",
        "    MOVE 8 TO WS-NUM.\n    DISPLAY WS-NUM.",
    ));
    assert_eq!(out, vec!["8"]);
}

#[test]
fn usage_comp_item_compiles() {
    let out = run_prints(&p(
        "01 WS-NUM PIC 9(3) USAGE COMPUTATIONAL VALUE 5.",
        "    MOVE 8 TO WS-NUM.\n    DISPLAY WS-NUM.",
    ));
    assert_eq!(out, vec!["8"]);
}

#[test]
fn sync_clause_compiles() {
    let out = run_prints(&p(
        "01 WS-REC.\n   05 WS-A PIC X(3).\n   05 WS-B PIC X(3) SYNC.",
        "    MOVE \"ABC\" TO WS-A.\n    MOVE \"XYZ\" TO WS-B.\n    DISPLAY WS-B.",
    ));
    assert_eq!(out, vec!["XYZ"]);
}

#[test]
fn occurs_depending_clause_compiles() {
    let out = run_prints(&p(
        "01 WS-TABLE PIC X(3) OCCURS 3 TIMES DEPENDING ON WS-COUNT.\n01 WS-COUNT PIC 9(1) VALUE 2.",
        "    MOVE \"A\" TO WS-TABLE(1).\n    MOVE \"B\" TO WS-TABLE(2).\n    DISPLAY WS-TABLE(1).\n    DISPLAY WS-TABLE(2).",
    ));
    assert_eq!(out, vec!["A", "B"]);
}

#[test]
fn value_then_move_changes_runtime() {
    let out = run_prints(&p(
        "01 WS-NUM PIC 9(3) VALUE 123.",
        "    DISPLAY WS-NUM.\n    MOVE 321 TO WS-NUM.\n    DISPLAY WS-NUM.",
    ));
    assert_eq!(out, vec!["123", "321"]);
}

#[test]
fn redefines_share_buffer_when_nesting_changes() {
    let out = run_prints(&p(
        "01 WS-BUF PIC X(2).\n01 WS-GRP REDEFINES WS-BUF.\n   05 WS-CH1 PIC X.\n   05 WS-CH2 PIC X.",
        "    MOVE \"AB\" TO WS-GRP.\n    MOVE WS-GRP TO WS-BUF.\n    DISPLAY WS-BUF.",
    ));
    assert_eq!(out, vec!["AB"]);
}
