use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn one_dimensional_occurs_table_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES.",
        "    MOVE 101 TO WS-ITEM(1).\n    MOVE 102 TO WS-ITEM(2).\n    MOVE 103 TO WS-ITEM(3).\n    DISPLAY WS-ITEM(1).\n    DISPLAY WS-ITEM(2).\n    DISPLAY WS-ITEM(3).",
    ));
    assert_eq!(out, vec!["101", "102", "103"]);
}

#[test]
fn indexed_table_declaration_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ITEM PIC X(4) OCCURS 5 TIMES INDEXED BY WS-IDX.",
        "    MOVE \"A\" TO WS-ITEM(1).\n    MOVE \"B\" TO WS-ITEM(2).\n    SET WS-IDX TO 1.\n    DISPLAY WS-ITEM(WS-IDX).\n    SET WS-IDX TO 2.\n    DISPLAY WS-ITEM(WS-IDX).",
    ));
    assert_eq!(out, vec!["A", "B"]);
}

#[test]
fn table_iteration_with_varying_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ITEM PIC 9(3) OCCURS 3 TIMES.\n01 WS-I PIC 9 VALUE 1.",
        "    MOVE 1 TO WS-ITEM(1).\n    MOVE 2 TO WS-ITEM(2).\n    MOVE 3 TO WS-ITEM(3).\n    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3\n        DISPLAY WS-ITEM(WS-I)\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn table_search_statement_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ENTRY OCCURS 10 TIMES INDEXED BY WS-IDX.\n      10 WS-KEY PIC X(4).",
        "    MOVE \"ABCD\" TO WS-KEY(1).\n    MOVE \"WXYZ\" TO WS-KEY(2).\n    MOVE \"LMNO\" TO WS-KEY(3).\n    SEARCH WS-ENTRY\n        WHEN WS-KEY(WS-IDX) = \"WXYZ\" DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND"]);
}

#[test]
fn table_search_not_found_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ENTRY OCCURS 3 TIMES ASCENDING KEY IS WS-KEY INDEXED BY WS-IDX.\n      10 WS-KEY PIC 9(3).",
        "    MOVE 10 TO WS-KEY(1).\n    MOVE 20 TO WS-KEY(2).\n    MOVE 30 TO WS-KEY(3).\n    SEARCH WS-ENTRY\n        AT END DISPLAY \"NONE\"\n        WHEN WS-KEY(WS-IDX) = 99 DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["NONE"]);
}

#[test]
fn table_search_all_statement_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ENTRY OCCURS 10 TIMES ASCENDING KEY IS WS-KEY INDEXED BY WS-IDX.\n      10 WS-KEY PIC 9(3).",
        "    MOVE 10 TO WS-KEY(1).\n    MOVE 20 TO WS-KEY(2).\n    MOVE 30 TO WS-KEY(3).\n    SEARCH ALL WS-ENTRY\n        WHEN WS-KEY(WS-IDX) = 20 DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND"]);
}

#[test]
fn table_search_all_not_found_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ENTRY OCCURS 3 TIMES ASCENDING KEY IS WS-KEY INDEXED BY WS-IDX.\n      10 WS-KEY PIC 9(3).",
        "    MOVE 10 TO WS-KEY(1).\n    MOVE 20 TO WS-KEY(2).\n    MOVE 30 TO WS-KEY(3).\n    SEARCH ALL WS-ENTRY\n        WHEN WS-KEY(WS-IDX) = 99 DISPLAY \"FOUND\"\n    END-SEARCH.\n    DISPLAY \"END\".",
    ));
    assert_eq!(out, vec!["END"]);
}

#[test]
fn table_nested_referential_runtime() {
    let out = run_prints(&p(
        "01 WS-COLS.\n   05 WS-ROW OCCURS 2 TIMES.\n      10 WS-COL PIC X(2) OCCURS 2 TIMES.",
        "    MOVE \"A1\" TO WS-COL(1,1).\n    MOVE \"A2\" TO WS-COL(1,2).\n    MOVE \"B1\" TO WS-COL(2,1).\n    MOVE \"B2\" TO WS-COL(2,2).\n    DISPLAY WS-COL(2,1).\n    DISPLAY WS-COL(1,2).",
    ));
    assert_eq!(out, vec!["B1", "A2"]);
}

#[test]
fn table_two_dimensional_access_runtime() {
    let out = run_prints(&p(
        "01 WS-MATRIX.\n   05 WS-ROW OCCURS 2 TIMES.\n      10 WS-COL PIC 9 OCCURS 3 TIMES VALUE 0.",
        "    MOVE 7 TO WS-COL(2,3).\n    DISPLAY WS-COL(2,3).",
    ));
    assert_eq!(out, vec!["7"]);
}

#[test]
fn table_indexed_element_display_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ITEM PIC X(5) OCCURS 2 TIMES.",
        "    MOVE \"HELLO\" TO WS-ITEM(1).\n    DISPLAY WS-ITEM(1).",
    ));
    assert_eq!(out, vec!["HELLO"]);
}
