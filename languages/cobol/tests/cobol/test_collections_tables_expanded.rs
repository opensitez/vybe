use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn one_dimensional_occurs_table_compiles() {
    compile_ok(&p(
        "01 WS-TABLE.\n   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES.",
        "    MOVE 123 TO WS-ITEM(1).",
    ));
}

#[test]
fn indexed_table_declaration_compiles() {
    compile_ok(&p(
        "01 WS-TABLE.\n   05 WS-ITEM PIC X(4) OCCURS 5 TIMES INDEXED BY WS-IDX.",
        "    SET WS-IDX TO 1.",
    ));
}

#[test]
fn table_iteration_with_varying_compiles() {
    compile_ok(&p(
        "01 WS-TABLE.\n   05 WS-ITEM PIC 9(3) OCCURS 3 TIMES.\n01 WS-I PIC 9 VALUE 1.",
        "    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3\n        MOVE WS-I TO WS-ITEM(WS-I)\n    END-PERFORM.",
    ));
}

#[test]
fn table_search_statement_compiles() {
    compile_ok(&p(
        "01 WS-TABLE.\n   05 WS-ENTRY OCCURS 10 TIMES INDEXED BY WS-IDX.\n      10 WS-KEY PIC X(4).",
        "    SEARCH WS-ENTRY\n        AT END DISPLAY \"NONE\"\n        WHEN WS-KEY(WS-IDX) = \"ABCD\" DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn table_search_all_statement_compiles() {
    compile_ok(&p(
        "01 WS-TABLE.\n   05 WS-ENTRY OCCURS 10 TIMES ASCENDING KEY IS WS-KEY INDEXED BY WS-IDX.\n      10 WS-KEY PIC 9(3).",
        "    SEARCH ALL WS-ENTRY\n        WHEN WS-KEY(WS-IDX) = 100 DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
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
