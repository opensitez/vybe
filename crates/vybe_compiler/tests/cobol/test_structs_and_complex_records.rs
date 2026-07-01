use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn nested_record_with_occurs_compiles() {
    compile_ok(&p(
        "01 WS-ORDER.\n   05 WS-ID PIC 9(6).\n   05 WS-LINES OCCURS 3 TIMES.\n      10 WS-SKU PIC X(10).\n      10 WS-QTY PIC 9(4).",
        "    MOVE 1 TO WS-ID.\n    MOVE \"SKU1\" TO WS-SKU(1).\n    MOVE 2 TO WS-QTY(1).",
    ));
}

#[test]
fn redefines_complex_record_compiles() {
    compile_ok(&p(
        "01 WS-BLOCK PIC X(30).\n01 WS-BLOCK-VIEW REDEFINES WS-BLOCK.\n   05 WS-CODE PIC X(5).\n   05 WS-AMOUNT PIC 9(5).\n   05 WS-DESC PIC X(20).",
        "    MOVE \"A1000\" TO WS-CODE.\n    MOVE 123 TO WS-AMOUNT.",
    ));
}

#[test]
fn move_corresponding_between_complex_records_compiles() {
    compile_ok(&p(
        "01 WS-SRC.\n   05 WS-NAME PIC X(10) VALUE \"ITEM\".\n   05 WS-COUNT PIC 9(4) VALUE 5.\n01 WS-DST.\n   05 WS-NAME PIC X(10).\n   05 WS-COUNT PIC 9(4).",
        "    MOVE CORRESPONDING WS-SRC TO WS-DST.",
    ));
}

#[test]
fn complex_record_initialize_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-A PIC X(5) VALUE \"AB\".\n   05 WS-B PIC 9(3) VALUE 9.\n   05 WS-C PIC X(5) VALUE \"CD\".",
        "    INITIALIZE WS-REC.",
    ));
}
