use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn occurs_set_and_read_element() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES INDEXED BY IX.",
        "    SET IX TO 3.\n    MOVE 77 TO E(IX).\n    DISPLAY E(IX).",
    ));
    assert_eq!(out, vec!["77"]);
}

#[test]
fn occurs_set_up_by() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES INDEXED BY IX.",
        "    SET IX TO 1.\n    SET IX UP BY 2.\n    MOVE 42 TO E(IX).\n    DISPLAY E(IX).",
    ));
    assert_eq!(out, vec!["42"]);
}

#[test]
fn occurs_set_down_by() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES INDEXED BY IX.",
        "    SET IX TO 5.\n    SET IX DOWN BY 3.\n    MOVE 11 TO E(IX).\n    DISPLAY E(IX).",
    ));
    assert_eq!(out, vec!["11"]);
}

#[test]
fn occurs_fill_via_index_loop() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9 OCCURS 5 TIMES INDEXED BY IX.",
        "    PERFORM VARYING IX FROM 1 BY 1 UNTIL IX > 5\n        MOVE IX TO E(IX)\n    END-PERFORM.\n    DISPLAY E(3).",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn occurs_sum_elements_via_index() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES INDEXED BY IX.\n01 S PIC 9(4) VALUE 0.",
        "    MOVE 10 TO E(1).\n    MOVE 20 TO E(2).\n    MOVE 30 TO E(3).\n    MOVE 40 TO E(4).\n    MOVE 50 TO E(5).\n    PERFORM VARYING IX FROM 1 BY 1 UNTIL IX > 5\n        ADD E(IX) TO S\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["0150"]);
}

#[test]
fn occurs_subscript_boundary_first() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X OCCURS 3 TIMES INDEXED BY IX.",
        "    MOVE \"A\" TO E(1).\n    DISPLAY E(1).",
    ));
    assert_eq!(out, vec!["A"]);
}

#[test]
fn occurs_subscript_boundary_last() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X OCCURS 3 TIMES INDEXED BY IX.",
        "    MOVE \"Z\" TO E(3).\n    DISPLAY E(3).",
    ));
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn occurs_two_indexes_two_tables() {
    let out = run_prints(&p(
        "01 T1.\n   05 A PIC 9 OCCURS 3 TIMES INDEXED BY IX1.\n01 T2.\n   05 B PIC 9 OCCURS 3 TIMES INDEXED BY IX2.",
        "    SET IX1 TO 1. SET IX2 TO 2.\n    MOVE 7 TO A(IX1).\n    MOVE 8 TO B(IX2).\n    DISPLAY A(IX1).\n    DISPLAY B(IX2).",
    ));
    assert_eq!(out, vec!["7", "8"]);
}

#[test]
fn occurs_indexed_in_search_compiles() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 10 TIMES INDEXED BY IX.",
        "    SEARCH E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = \"ABC\"\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn occurs_index_set_then_move_multiple() {
    let out = run_prints(&p(
        "01 T.\n   05 VAL PIC 9(2) OCCURS 4 TIMES INDEXED BY IDX.",
        "    SET IDX TO 1.\n    MOVE 11 TO VAL(IDX).\n    SET IDX UP BY 1.\n    MOVE 22 TO VAL(IDX).\n    SET IDX UP BY 1.\n    MOVE 33 TO VAL(IDX).\n    DISPLAY VAL(1).\n    DISPLAY VAL(2).\n    DISPLAY VAL(3).",
    ));
    assert_eq!(out, vec!["11", "22", "33"]);
}

#[test]
fn occurs_two_dim_inner_loop() {
    compile_ok(&p(
        "01 MATRIX.\n   05 ROW OCCURS 3 TIMES INDEXED BY RI.\n      10 COL PIC 9 OCCURS 3 TIMES INDEXED BY CI.",
        "    PERFORM VARYING RI FROM 1 BY 1 UNTIL RI > 3\n        PERFORM VARYING CI FROM 1 BY 1 UNTIL CI > 3\n            MOVE 0 TO COL(RI CI)\n        END-PERFORM\n    END-PERFORM.",
    ));
}

#[test]
fn occurs_indexed_by_no_redefine() {
    compile_ok(&p(
        "01 DATA-TABLE.\n   05 DT-ENTRY PIC X(10) OCCURS 20 TIMES INDEXED BY DT-IDX.",
        "    SET DT-IDX TO 1.\n    MOVE \"FIRST\" TO DT-ENTRY(DT-IDX).",
    ));
}

#[test]
fn occurs_set_to_integer_variable() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9 OCCURS 5 TIMES INDEXED BY IX.\n01 N PIC 9 VALUE 4.",
        "    SET IX TO N.\n    MOVE 9 TO E(IX).\n    DISPLAY E(4).",
    ));
    assert_eq!(out, vec!["9"]);
}

#[test]
fn occurs_set_from_index_to_integer() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X OCCURS 5 TIMES INDEXED BY IX.\n01 POS PIC 9 VALUE 0.",
        "    SET IX TO 3.\n    SET POS TO IX.",
    ));
}

#[test]
fn occurs_perform_down_decrement() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 4 TIMES INDEXED BY IX.\n01 S PIC 9(4) VALUE 0.",
        "    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3). MOVE 40 TO E(4).\n    SET IX TO 4.\n    PERFORM UNTIL IX < 1\n        ADD E(IX) TO S\n        SET IX DOWN BY 1\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["0100"]);
}

#[test]
fn occurs_search_all_requires_ascending_key_compiles() {
    compile_ok(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 10 TIMES ASCENDING KEY E INDEXED BY IX.",
        "    SEARCH ALL E\n        AT END DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = 5\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn occurs_copy_table_element_to_variable() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(5) OCCURS 3 TIMES INDEXED BY IX.\n01 COPY PIC X(5) VALUE SPACES.",
        "    MOVE \"HELLO\" TO E(2).\n    SET IX TO 2.\n    MOVE E(IX) TO COPY.\n    DISPLAY COPY.",
    ));
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn occurs_loop_counts_greater_than_value() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 10 TIMES INDEXED BY IX.\n01 CNT PIC 9(2) VALUE 0.",
        "    PERFORM VARYING IX FROM 1 BY 1 UNTIL IX > 10\n        MOVE IX TO E(IX)\n    END-PERFORM.\n    PERFORM VARYING IX FROM 1 BY 1 UNTIL IX > 10\n        IF E(IX) > 5\n            ADD 1 TO CNT\n        END-IF\n    END-PERFORM.\n    DISPLAY CNT.",
    ));
    // Elements 6,7,8,9,10 are > 5
    assert_eq!(out, vec!["05"]);
}

#[test]
fn occurs_index_preserved_after_loop() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9 OCCURS 5 TIMES INDEXED BY IX.",
        "    PERFORM VARYING IX FROM 1 BY 1 UNTIL IX > 5\n        CONTINUE\n    END-PERFORM.\n    DISPLAY IX.",
    ));
    // After loop IX = 6
    assert_eq!(out, vec!["6"]);
}

#[test]
fn occurs_with_depending_on_compiles() {
    compile_ok(&p(
        "01 MAX-ITEMS PIC 9(3) VALUE 10.\n01 T.\n   05 E PIC X(5) OCCURS 1 TO 50 TIMES DEPENDING ON MAX-ITEMS INDEXED BY IX.",
        "    SET IX TO 1.\n    MOVE \"FIRST\" TO E(IX).",
    ));
}

#[test]
fn occurs_complex_key_field_compiles() {
    compile_ok(&p(
        "01 LOOKUP.\n   05 ENTRY OCCURS 20 TIMES ASCENDING KEY CODE-VAL INDEXED BY LK-IDX.\n      10 CODE-VAL PIC X(4).\n      10 DESC-VAL PIC X(20).",
        "    SET LK-IDX TO 1.\n    MOVE \"AAAA\" TO CODE-VAL(LK-IDX).",
    ));
}

#[test]
fn occurs_negative_direction_set_and_read() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE 55 TO E(3).\n    SET IX TO 5.\n    SET IX DOWN BY 2.\n    DISPLAY E(IX).",
    ));
    assert_eq!(out, vec!["55"]);
}

#[test]
fn occurs_table_filled_displayed_first() {
    let out = run_prints(&p(
        "01 T.\n   05 SCORE PIC 9(3) OCCURS 5 TIMES INDEXED BY IDX.",
        "    MOVE 100 TO SCORE(1).\n    MOVE 200 TO SCORE(2).\n    MOVE 300 TO SCORE(3).\n    MOVE 400 TO SCORE(4).\n    MOVE 500 TO SCORE(5).\n    DISPLAY SCORE(1).",
    ));
    assert_eq!(out, vec!["100"]);
}

#[test]
fn occurs_search_linear_at_end() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO E(1).\n    MOVE \"BBB\" TO E(2).\n    MOVE \"CCC\" TO E(3).\n    MOVE \"DDD\" TO E(4).\n    MOVE \"EEE\" TO E(5).\n    SEARCH E\n        AT END\n            DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = \"ZZZ\"\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["NOT FOUND"]);
}

#[test]
fn occurs_search_linear_found() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE \"AAA\" TO E(1).\n    MOVE \"BBB\" TO E(2).\n    MOVE \"CCC\" TO E(3).\n    SET IX TO 1.\n    SEARCH E\n        AT END\n            DISPLAY \"NOT FOUND\"\n        WHEN E(IX) = \"BBB\"\n            DISPLAY \"FOUND BBB\"\n    END-SEARCH.",
    ));
    assert_eq!(out, vec!["FOUND BBB"]);
}

#[test]
fn occurs_display_all_five_elements() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9 OCCURS 5 TIMES INDEXED BY IX.",
        "    MOVE 1 TO E(1). MOVE 2 TO E(2). MOVE 3 TO E(3). MOVE 4 TO E(4). MOVE 5 TO E(5).\n    PERFORM VARYING IX FROM 1 BY 1 UNTIL IX > 5\n        DISPLAY E(IX)\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["1", "2", "3", "4", "5"]);
}
