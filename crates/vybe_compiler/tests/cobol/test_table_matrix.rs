use super::helpers::compile_ok_check;

fn make_table_program(occurs: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-TABLE.\n   05 WS-ITEM PIC 9(3) OCCURS {occurs}.\nPROCEDURE DIVISION.\n{body}\n    STOP RUN."
    )
}

#[test]
fn test_table_indexing_matrix() {
    let programs = [
        make_table_program("3 TIMES", "    MOVE 1 TO WS-ITEM(1)."),
        make_table_program("5 TIMES", "    MOVE 2 TO WS-ITEM(2)."),
        make_table_program("7 TIMES", "    MOVE 3 TO WS-ITEM(3)."),
        make_table_program("9 TIMES", "    MOVE 4 TO WS-ITEM(4)."),
        make_table_program("10 TIMES", "    MOVE 5 TO WS-ITEM(5)."),
        make_table_program("12 TIMES", "    MOVE 6 TO WS-ITEM(6)."),
        make_table_program("15 TIMES", "    MOVE 7 TO WS-ITEM(7)."),
    ];

    for program in programs {
        assert!(compile_ok_check(&program), "table case failed:\n{program}");
    }
}
