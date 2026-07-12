use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn indexed_array_assignments_compile() {
    compile_ok(&p(
        "01 WS-TABLE PIC 9(3) OCCURS 3 TIMES.\n01 WS-INDEX PIC 9(1) VALUE 1.",
        "    MOVE 10 TO WS-TABLE(1).\n    MOVE 20 TO WS-TABLE(2).\n    MOVE 30 TO WS-TABLE(3).",
    ));
}

#[test]
fn array_indexed_loop_compile() {
    compile_ok(&p(
        "01 WS-TABLE PIC 9(3) OCCURS 3 TIMES.\n01 WS-INDEX PIC 9(1) VALUE 1.",
        "    PERFORM VARYING WS-INDEX FROM 1 BY 1 UNTIL WS-INDEX > 3\n        MOVE WS-INDEX TO WS-TABLE(WS-INDEX)\n    END-PERFORM.",
    ));
}

#[test]
fn array_indexed_values_runtime_display_expected_cells() {
    let output = run_prints(&p(
        "01 WS-TABLE PIC 9 OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.",
        "    MOVE 4 TO WS-TABLE(1).\n    MOVE 5 TO WS-TABLE(2).\n    MOVE 6 TO WS-TABLE(3).\n    DISPLAY WS-TABLE(1).\n    DISPLAY WS-TABLE(2).\n    DISPLAY WS-TABLE(3).",
    ));
    assert_eq!(output, vec!["4", "5", "6"]);
}

#[test]
fn array_varying_loop_runtime_populates_sequence() {
    let output = run_prints(&p(
        "01 WS-TABLE PIC 9 OCCURS 3 TIMES.\n01 I PIC 9.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        MOVE I TO WS-TABLE(I)\n    END-PERFORM.\n    DISPLAY WS-TABLE(1).\n    DISPLAY WS-TABLE(2).\n    DISPLAY WS-TABLE(3).",
    ));
    assert_eq!(output, vec!["1", "2", "3"]);
}

#[test]
fn array_runtime_sum_over_occurs_table() {
    let output = run_prints(&p(
        "01 T PIC 9 OCCURS 3 TIMES.\n01 I PIC 9 VALUE 0.\n01 S PIC 99 VALUE 0.",
        "    MOVE 2 TO T(1).\n    MOVE 3 TO T(2).\n    MOVE 4 TO T(3).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        ADD T(I) TO S\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(output, vec!["9"]);
}

#[test]
fn array_runtime_overwrite_single_cell_only() {
    let output = run_prints(&p(
        "01 T PIC 9 OCCURS 3 TIMES.\n01 I PIC 9 VALUE 0.",
        "    MOVE 1 TO T(1).\n    MOVE 2 TO T(2).\n    MOVE 3 TO T(3).\n    MOVE 9 TO T(2).\n    DISPLAY T(1).\n    DISPLAY T(2).\n    DISPLAY T(3).",
    ));
    assert_eq!(output, vec!["1", "9", "3"]);
}
