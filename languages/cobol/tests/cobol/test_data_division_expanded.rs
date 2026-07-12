use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn numeric_item_with_value_compiles() {
    compile_ok(&p("01 WS-NUM PIC 9(3) VALUE 5.", "    MOVE 7 TO WS-NUM."));
}
#[test]
fn signed_numeric_item_compiles() {
    compile_ok(&p("01 WS-NUM PIC S9(3) VALUE -3.", "    MOVE 4 TO WS-NUM."));
}
#[test]
fn edited_numeric_item_compiles() {
    compile_ok(&p(
        "01 WS-ED PIC ZZ9.99 VALUE 12.34.",
        "    MOVE 7.5 TO WS-ED.",
    ));
}
#[test]
fn alphanumeric_item_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(12) VALUE \"HELLO\".",
        "    MOVE \"WORLD\" TO WS-TXT.",
    ));
}
#[test]
fn alphanumeric_editted_item_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(12) VALUE \"ABC\".",
        "    MOVE \"123\" TO WS-TXT.",
    ));
}
#[test]
fn group_item_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-A PIC X(3).\n   05 WS-B PIC 9(2).",
        "    MOVE \"XYZ\" TO WS-A.",
    ));
}
#[test]
fn nested_group_item_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-INFO.\n      10 WS-NAME PIC X(5).\n      10 WS-AGE PIC 9(2).",
        "    MOVE \"BOB\" TO WS-NAME.",
    ));
}
#[test]
fn occurs_clause_compiles() {
    compile_ok(&p(
        "01 WS-TABLE PIC X(3) OCCURS 5 TIMES.",
        "    MOVE \"A\" TO WS-TABLE(1).",
    ));
}
#[test]
fn occurs_with_index_compiles() {
    compile_ok(&p(
        "01 WS-TABLE PIC 9(2) OCCURS 4 TIMES.\n01 WS-IDX PIC 9(1) VALUE 1.",
        "    MOVE 10 TO WS-TABLE(WS-IDX).",
    ));
}
#[test]
fn redefines_item_compiles() {
    compile_ok(&p(
        "01 WS-BUF PIC X(10).\n01 WS-NUM REDEFINES WS-BUF PIC 9(10).",
        "    MOVE 42 TO WS-NUM.",
    ));
}
#[test]
fn redefines_group_item_compiles() {
    compile_ok(&p(
        "01 WS-BUF PIC X(20).\n01 WS-GRP REDEFINES WS-BUF.\n   05 WS-A PIC X(4).\n   05 WS-B PIC 9(2).",
        "    MOVE \"HI\" TO WS-A.",
    ));
}
#[test]
fn value_clause_string_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(5) VALUE \"A\".",
        "    MOVE \"B\" TO WS-TXT.",
    ));
}
#[test]
fn value_clause_numeric_compiles() {
    compile_ok(&p(
        "01 WS-NUM PIC 9(3) VALUE 100.",
        "    MOVE 200 TO WS-NUM.",
    ));
}
#[test]
fn value_clause_spaces_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(5) VALUE SPACES.",
        "    MOVE \"X\" TO WS-TXT.",
    ));
}
#[test]
fn value_clause_zeros_compiles() {
    compile_ok(&p(
        "01 WS-NUM PIC 9(5) VALUE ZEROS.",
        "    MOVE 1 TO WS-NUM.",
    ));
}
#[test]
fn blank_when_zero_clause_compiles() {
    compile_ok(&p(
        "01 WS-NUM PIC 9(3) BLANK WHEN ZERO VALUE 0.",
        "    MOVE 0 TO WS-NUM.",
    ));
}
#[test]
fn just_justified_right_clause_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(5) JUSTIFIED RIGHT VALUE \"A\".",
        "    MOVE \"B\" TO WS-TXT.",
    ));
}
#[test]
fn usage_display_item_compiles() {
    compile_ok(&p(
        "01 WS-NUM PIC 9(3) USAGE DISPLAY VALUE 5.",
        "    MOVE 8 TO WS-NUM.",
    ));
}
#[test]
fn usage_comp_item_compiles() {
    compile_ok(&p(
        "01 WS-NUM PIC 9(3) USAGE COMPUTATIONAL VALUE 5.",
        "    MOVE 8 TO WS-NUM.",
    ));
}
#[test]
fn sync_clause_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-A PIC X(3).\n   05 WS-B PIC X(3) SYNC.",
        "    MOVE \"ABC\" TO WS-A.",
    ));
}
#[test]
fn occurs_depending_clause_compiles() {
    compile_ok(&p(
        "01 WS-TABLE PIC X(3) OCCURS 3 TIMES DEPENDING ON WS-COUNT.\n01 WS-COUNT PIC 9(1) VALUE 2.",
        "    MOVE \"A\" TO WS-TABLE(1).",
    ));
}
