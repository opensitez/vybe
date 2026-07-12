use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn perform_varying_table_traversal_runtime() {
    let out = run_prints(&p(
        "01 WS-TABLE.\n   05 WS-ITEM PIC 9(2) OCCURS 3 TIMES.\n01 WS-I PIC 9 VALUE 1.",
        "    MOVE 11 TO WS-ITEM(1).\n    MOVE 22 TO WS-ITEM(2).\n    MOVE 33 TO WS-ITEM(3).\n    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3\n        DISPLAY WS-ITEM(WS-I)\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["11", "22", "33"]);
}

#[test]
fn search_table_iterator_style_compiles() {
    compile_ok(&p(
        "01 WS-TABLE.\n   05 WS-ENTRY OCCURS 5 TIMES INDEXED BY WS-IDX.\n      10 WS-KEY PIC X(3).",
        "    SEARCH WS-ENTRY\n        AT END DISPLAY \"NONE\"\n        WHEN WS-KEY(WS-IDX) = \"ABC\" DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn search_all_table_iterator_style_compiles() {
    compile_ok(&p(
        "01 WS-TABLE.\n   05 WS-ENTRY OCCURS 5 TIMES ASCENDING KEY IS WS-KEY INDEXED BY WS-IDX.\n      10 WS-KEY PIC 9(3).",
        "    SEARCH ALL WS-ENTRY\n        WHEN WS-KEY(WS-IDX) = 200 DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn perform_until_iterator_counter_runtime() {
    let out = run_prints(&p(
        "01 WS-I PIC 9 VALUE 0.",
        "    PERFORM UNTIL WS-I >= 3\n        ADD 1 TO WS-I\n        DISPLAY WS-I\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["1", "2", "3"]);
}
