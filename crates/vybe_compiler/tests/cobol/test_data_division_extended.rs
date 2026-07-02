use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn data_division_numeric_picture_compiles() {
    compile_ok(&p("01 WS-N PIC 9(5) VALUE 12345.", "    DISPLAY WS-N."));
}

#[test]
fn data_division_alphanumeric_picture_compiles() {
    compile_ok(&p("01 WS-S PIC X(8) VALUE \"HELLO\".", "    DISPLAY WS-S."));
}

#[test]
fn data_division_group_item_compiles() {
    compile_ok(&p(
        "01 WS-GRP.\n   05 WS-A PIC X(2) VALUE \"AB\".\n   05 WS-B PIC X(2) VALUE \"CD\".",
        "    DISPLAY WS-GRP.",
    ));
}

#[test]
fn data_division_level_77_item_compiles() {
    compile_ok(&p("77 WS-VAL PIC 9(3) VALUE 100.", "    DISPLAY WS-VAL."));
}

#[test]
fn data_division_occurs_clause_compiles() {
    compile_ok(&p(
        "01 WS-TBL.\n   05 WS-ITEM PIC 9(2) OCCURS 3 TIMES.",
        "    DISPLAY WS-ITEM(1).",
    ));
}

#[test]
fn data_division_redefines_clause_compiles() {
    compile_ok(&p(
        "01 WS-BUF PIC X(4) VALUE \"ABCD\".\n01 WS-NUM REDEFINES WS-BUF PIC 9(4).",
        "    DISPLAY WS-NUM.",
    ));
}

#[test]
fn data_division_value_clause_in_group_item_compiles() {
    compile_ok(&p(
        "01 WS-GRP.\n   05 WS-A PIC X(2) VALUE \"AA\".\n   05 WS-B PIC X(2) VALUE \"BB\".",
        "    DISPLAY WS-A.\n    DISPLAY WS-B.",
    ));
}

#[test]
fn data_division_occurs_with_index_compiles() {
    compile_ok(&p(
        "01 WS-TBL.\n   05 WS-ITEM PIC 9(2) OCCURS 2 TIMES INDEXED BY WS-IDX.",
        "    DISPLAY WS-ITEM(1).",
    ));
}
