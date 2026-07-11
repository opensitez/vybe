use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn occurs_fixed_table_compiles() {
    compile_ok(&p(
        "01 TBL PIC 9(2) OCCURS 5 TIMES.",
        "    MOVE 10 TO TBL(1).",
    ));
}
#[test]
fn occurs_group_table_compiles() {
    compile_ok(&p(
        "01 TBL.\n   05 ITM OCCURS 3 TIMES.\n      10 V PIC X(3).",
        "    MOVE \"A\" TO V(1).",
    ));
}
#[test]
fn occurs_indexed_compiles() {
    compile_ok(&p(
        "01 TBL PIC 9(2) OCCURS 5 TIMES INDEXED BY IDX.",
        "    SET IDX TO 1.",
    ));
}
#[test]
fn occurs_depending_compiles() {
    compile_ok(&p(
        "01 CNT PIC 9 VALUE 2.\n01 TBL PIC X(2) OCCURS 1 TO 5 TIMES DEPENDING ON CNT.",
        "    MOVE \"AA\" TO TBL(1).",
    ));
}
#[test]
fn table_set_up_down_compiles() {
    compile_ok(&p(
        "01 TBL PIC 9(2) OCCURS 5 TIMES INDEXED BY IDX.",
        "    SET IDX TO 1.\n    SET IDX UP BY 1.\n    SET IDX DOWN BY 1.",
    ));
}
#[test]
fn table_search_compiles() {
    compile_ok(&p(
        "01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).",
        "    SEARCH E\n        WHEN K(I) = \"AAA\" DISPLAY \"F\"\n    END-SEARCH.",
    ));
}
#[test]
fn table_search_all_compiles() {
    compile_ok(&p(
        "01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(3).",
        "    SEARCH ALL E\n        WHEN K(I) = 200 DISPLAY \"F\"\n    END-SEARCH.",
    ));
}
#[test]
fn table_copy_varying_compiles() {
    compile_ok(&p(
        "01 A PIC X(2) OCCURS 3 TIMES.\n01 B PIC X(2) OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        MOVE A(I) TO B(I)\n    END-PERFORM.",
    ));
}
#[test]
fn table_runtime_display_values() {
    let o = run_prints(&p(
        "01 T PIC 9 OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.",
        "    MOVE 1 TO T(1). MOVE 2 TO T(2). MOVE 3 TO T(3).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        DISPLAY T(I)\n    END-PERFORM.",
    ));
    assert_eq!(o, vec!["1", "2", "3"]);
}
#[test]
fn two_dimensional_table_compiles() {
    compile_ok(&p(
        "01 T.\n   05 R OCCURS 2 TIMES.\n      10 C PIC 9 OCCURS 2 TIMES.",
        "    MOVE 5 TO C(2,2).",
    ));
}
#[test]
fn table_move_zeros_compiles() {
    compile_ok(&p(
        "01 T PIC 9(2) OCCURS 4 TIMES VALUE 12.",
        "    MOVE ZEROS TO T.",
    ));
}
#[test]
fn table_group_move_compiles() {
    compile_ok(&p(
        "01 G1.\n   05 A PIC X(2) OCCURS 2 TIMES.\n01 G2.\n   05 B PIC X(2) OCCURS 2 TIMES.",
        "    MOVE G1 TO G2.",
    ));
}
#[test]
fn table_display_element_compiles() {
    compile_ok(&p("01 T PIC X(3) OCCURS 2 TIMES.", "    DISPLAY T(1)."));
}
#[test]
fn table_compute_element_compiles() {
    compile_ok(&p(
        "01 T PIC 9(3) OCCURS 2 TIMES.\n01 X PIC 9(3) VALUE 4.",
        "    COMPUTE T(1) = X + 2.",
    ));
}
#[test]
fn table_nested_occurs_compiles() {
    compile_ok(&p(
        "01 T.\n   05 O1 OCCURS 2 TIMES.\n      10 O2 OCCURS 2 TIMES.\n         15 V PIC X(2).",
        "    MOVE \"AA\" TO V(1,1).",
    ));
}
#[test]
fn table_index_var_mix_compiles() {
    compile_ok(&p(
        "01 T PIC 9(2) OCCURS 3 TIMES.\n01 I PIC 9 VALUE 2.",
        "    MOVE 9 TO T(I).",
    ));
}
#[test]
fn table_if_condition_compiles() {
    compile_ok(&p(
        "01 T PIC 9(2) OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.",
        "    IF T(I) = 0 DISPLAY \"Z\" END-IF.",
    ));
}
#[test]
fn table_evaluate_condition_compiles() {
    compile_ok(&p(
        "01 T PIC 9 OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.",
        "    EVALUATE T(I)\n        WHEN 1 DISPLAY \"O\"\n        WHEN OTHER DISPLAY \"X\"\n    END-EVALUATE.",
    ));
}
