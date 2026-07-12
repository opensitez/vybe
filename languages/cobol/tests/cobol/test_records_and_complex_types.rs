use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn nested_group_item_compiles() {
    compile_ok(&p(
        "01 WS-REC.\n   05 WS-NAME PIC X(10).\n   05 WS-AGE PIC 9(2).",
        "    MOVE \"ALICE\" TO WS-NAME.\n    MOVE 30 TO WS-AGE.",
    ));
}

#[test]
fn group_move_corresponding_compiles() {
    compile_ok(&p(
        "01 WS-SRC.\n   05 WS-NAME PIC X(10) VALUE \"BOB\".\n   05 WS-AGE PIC 9(2) VALUE 41.\n01 WS-DST.\n   05 WS-NAME PIC X(10).\n   05 WS-AGE PIC 9(2).",
        "    MOVE CORRESPONDING WS-SRC TO WS-DST.",
    ));
}

#[test]
fn redefines_group_item_compiles() {
    compile_ok(&p(
        "01 WS-BUFFER PIC X(20).\n01 WS-FIELD REDEFINES WS-BUFFER.\n   05 WS-CHAR PIC X(1) OCCURS 20 TIMES.",
        "    MOVE \"A\" TO WS-CHAR(1).",
    ));
}
