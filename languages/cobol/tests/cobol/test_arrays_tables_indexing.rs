use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn occurs_fixed_table_compiles() {
    let o = run_prints(&p(
        "01 TBL PIC 9(2) OCCURS 5 TIMES.",
        "    MOVE 10 TO TBL(1).\n    MOVE 20 TO TBL(2).\n    DISPLAY TBL(1).\n    DISPLAY TBL(2).",
    ));
    assert_eq!(o, vec!["10", "20"]);
}
#[test]
fn occurs_group_table_compiles() {
    let o = run_prints(&p(
        "01 TBL.\n   05 ITM OCCURS 3 TIMES.\n      10 V PIC X(3).",
        "    MOVE \"AAA\" TO V(1).\n    MOVE \"BBB\" TO V(2).\n    DISPLAY V(1).\n    DISPLAY V(2).",
    ));
    assert_eq!(o, vec!["AAA", "BBB"]);
}
#[test]
fn occurs_indexed_compiles() {
    let o = run_prints(&p(
        "01 TBL PIC 9(2) OCCURS 5 TIMES INDEXED BY IDX.",
        "    SET IDX TO 1.\n    MOVE 11 TO TBL(IDX).\n    SET IDX UP BY 1.\n    MOVE 22 TO TBL(IDX).\n    SET IDX DOWN BY 1.\n    MOVE 33 TO TBL(IDX).\n    DISPLAY TBL(1).\n    DISPLAY TBL(2).",
    ));
    assert_eq!(o, vec!["33", "22"]);
}
#[test]
fn occurs_depending_compiles() {
    let o = run_prints(&p(
        "01 CNT PIC 9 VALUE 2.\n01 TBL PIC X(2) OCCURS 1 TO 5 TIMES DEPENDING ON CNT.",
        "    MOVE 2 TO CNT.\n    MOVE \"AA\" TO TBL(1).\n    MOVE \"BB\" TO TBL(2).\n    DISPLAY TBL(1).\n    DISPLAY TBL(2).",
    ));
    assert_eq!(o, vec!["AA", "BB"]);
}
#[test]
fn table_set_up_down_compiles() {
    let o = run_prints(&p(
        "01 TBL PIC 9(2) OCCURS 5 TIMES INDEXED BY IDX.",
        "    SET IDX TO 1.\n    MOVE 10 TO TBL(IDX).\n    SET IDX UP BY 1.\n    MOVE 20 TO TBL(IDX).\n    SET IDX DOWN BY 1.\n    DISPLAY TBL(IDX).\n    DISPLAY TBL(2).",
    ));
    assert_eq!(o, vec!["10", "20"]);
}
#[test]
fn table_search_compiles() {
    let o = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).",
        "    MOVE \"AAA\" TO K(1).\n    MOVE \"BBB\" TO K(2).\n    MOVE \"CCC\" TO K(3).\n    SEARCH E\n        WHEN K(I) = \"BBB\" DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(o, vec!["FOUND"]);
}
#[test]
fn table_search_all_compiles() {
    let o = run_prints(&p(
        "01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(3).",
        "    MOVE 10 TO K(1).\n    MOVE 20 TO K(2).\n    MOVE 30 TO K(3).\n    SEARCH ALL E\n        WHEN K(I) = 20 DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(o, vec!["FOUND"]);
}
#[test]
fn table_copy_varying_compiles() {
    let o = run_prints(&p(
        "01 A PIC X(2) OCCURS 3 TIMES.\n01 B PIC X(2) OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.",
        "    MOVE \"AA\" TO A(1).\n    MOVE \"BB\" TO A(2).\n    MOVE \"CC\" TO A(3).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        MOVE A(I) TO B(I)\n    END-PERFORM.\n    DISPLAY B(1).\n    DISPLAY B(2).\n    DISPLAY B(3).",
    ));
    assert_eq!(o, vec!["AA", "BB", "CC"]);
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
    let o = run_prints(&p(
        "01 T.\n   05 R OCCURS 2 TIMES.\n      10 C PIC 9 OCCURS 2 TIMES.",
        "    MOVE 5 TO C(1,1).\n    MOVE 6 TO C(1,2).\n    MOVE 7 TO C(2,1).\n    MOVE 8 TO C(2,2).\n    DISPLAY C(2,2).\n    DISPLAY C(1,2).",
    ));
    assert_eq!(o, vec!["8", "6"]);
}
#[test]
fn table_move_zeros_compiles() {
    let o = run_prints(&p(
        "01 T PIC 9(2) OCCURS 4 TIMES VALUE 12.",
        "    MOVE ZEROS TO T.\n    DISPLAY T(1).\n    DISPLAY T(2).\n    DISPLAY T(3).\n    DISPLAY T(4).",
    ));
    assert_eq!(o, vec!["00", "00", "00", "00"]);
}
#[test]
fn table_group_move_compiles() {
    let o = run_prints(&p(
        "01 G1.\n   05 A PIC X(2) OCCURS 2 TIMES.\n01 G2.\n   05 B PIC X(2) OCCURS 2 TIMES.",
        "    MOVE \"AA\" TO A(1).\n    MOVE \"BB\" TO A(2).\n    MOVE G1 TO G2.\n    DISPLAY B(1).\n    DISPLAY B(2).",
    ));
    assert_eq!(o, vec!["AA", "BB"]);
}
#[test]
fn table_display_element_compiles() {
    let o = run_prints(&p(
        "01 T PIC X(3) OCCURS 2 TIMES.",
        "    MOVE \"ABC\" TO T(1).\n    DISPLAY T(1).",
    ));
    assert_eq!(o, vec!["ABC"]);
}
#[test]
fn table_compute_element_compiles() {
    let o = run_prints(&p(
        "01 T PIC 9(3) OCCURS 2 TIMES.\n01 X PIC 9(3) VALUE 4.",
        "    COMPUTE T(1) = X + 2.\n    DISPLAY T(1).",
    ));
    assert_eq!(o, vec!["6"]);
}
#[test]
fn table_nested_occurs_compiles() {
    let o = run_prints(&p(
        "01 T.\n   05 O1 OCCURS 2 TIMES.\n      10 O2 OCCURS 2 TIMES.\n         15 V PIC X(2).",
        "    MOVE \"AA\" TO V(1,1).\n    MOVE \"BB\" TO V(2,2).\n    DISPLAY V(1,1).\n    DISPLAY V(2,2).",
    ));
    assert_eq!(o, vec!["AA", "BB"]);
}

#[test]
fn table_search_found_runtime() {
    let o = run_prints(&p(
        "01 T.\n   05 E OCCURS 3 TIMES ASCENDING KEY IS K INDEXED BY I.\n      10 K PIC 9(2).",
        "    MOVE 10 TO K(1).\n    MOVE 20 TO K(2).\n    MOVE 30 TO K(3).\n    SEARCH ALL E\n        WHEN K(I) = 20\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(o, vec!["FOUND"]);
}
#[test]
fn table_index_var_mix_compiles() {
    let o = run_prints(&p(
        "01 T PIC 9(2) OCCURS 3 TIMES.\n01 I PIC 9 VALUE 2.",
        "    MOVE 9 TO T(I).\n    DISPLAY T(I).",
    ));
    assert_eq!(o, vec!["9"]);
}
#[test]
fn table_if_condition_compiles() {
    let o = run_prints(&p(
        "01 T PIC 9(2) OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.",
        "    MOVE 2 TO I.\n    MOVE 5 TO T(2).\n    IF T(I) = 5 DISPLAY \"MATCH\" END-IF.",
    ));
    assert_eq!(o, vec!["MATCH"]);
}
#[test]
fn table_evaluate_condition_compiles() {
    let o = run_prints(&p(
        "01 T PIC 9 OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.",
        "    MOVE 2 TO I.\n    MOVE 2 TO T(2).\n    EVALUATE T(I)\n        WHEN 1 DISPLAY \"ONE\"\n        WHEN 2 DISPLAY \"TWO\"\n        WHEN OTHER DISPLAY \"X\"\n    END-EVALUATE.",
    ));
    assert_eq!(o, vec!["TWO"]);
}

#[test]
fn table_search_not_found_runtime() {
    let o = run_prints(&p(
        "01 T.\n   05 E OCCURS 4 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(2).",
        "    MOVE 11 TO K(1).\n    MOVE 22 TO K(2).\n    MOVE 33 TO K(3).\n    MOVE 44 TO K(4).\n    SEARCH E\n        AT END DISPLAY \"NOT-FOUND\"\n        WHEN K(I) = 99\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
    assert_eq!(o, vec!["NOT-FOUND"]);
}

#[test]
fn table_search_all_not_found_runtime() {
    let o = run_prints(&p(
        "01 T.\n   05 E OCCURS 4 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(2).",
        "    MOVE 11 TO K(1).\n    MOVE 22 TO K(2).\n    MOVE 33 TO K(3).\n    MOVE 44 TO K(4).\n    SEARCH ALL E\n        WHEN K(I) = 99 DISPLAY \"FOUND\"\n    END-SEARCH.\n    DISPLAY \"DONE\".",
    ));
    assert_eq!(o, vec!["DONE"]);
}

#[test]
fn table_copy_occurrence_runtime() {
    let o = run_prints(&p(
        "01 SRC.\n   05 V PIC X(2) OCCURS 3 TIMES.\n01 DST.\n   05 W PIC X(2) OCCURS 3 TIMES.",
        "    MOVE \"AA\" TO V(1).\n    MOVE \"BB\" TO V(2).\n    MOVE \"CC\" TO V(3).\n    MOVE V(1) TO W(3).\n    MOVE V(2) TO W(1).\n    MOVE V(3) TO W(2).\n    DISPLAY W(1).\n    DISPLAY W(2).\n    DISPLAY W(3).",
    ));
    assert_eq!(o, vec!["BB", "CC", "AA"]);
}

#[test]
fn table_varying_sum_runtime() {
    let o = run_prints(&p(
        "01 T PIC 9(2) OCCURS 4 TIMES.\n01 I PIC 9 VALUE 1.\n01 TOT PIC 9(3) VALUE 0.",
        "    MOVE 1 TO T(1).\n    MOVE 2 TO T(2).\n    MOVE 3 TO T(3).\n    MOVE 4 TO T(4).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 4\n        ADD T(I) TO TOT\n    END-PERFORM.\n    DISPLAY TOT.",
    ));
    assert_eq!(o, vec!["10"]);
}

#[test]
fn table_index_set_up_down_runtime() {
    let o = run_prints(&p(
        "01 T PIC X(1) OCCURS 5 TIMES INDEXED BY IDX.\n",
        "    MOVE \"A\" TO T(1).\n    MOVE \"B\" TO T(2).\n    MOVE \"C\" TO T(3).\n    MOVE \"D\" TO T(4).\n    MOVE \"E\" TO T(5).\n    SET IDX TO 4.\n    DISPLAY T(IDX).\n    SET IDX DOWN BY 2.\n    DISPLAY T(IDX).\n    SET IDX UP BY 1.\n    DISPLAY T(IDX).",
    ));
    assert_eq!(o, vec!["D", "B", "C"]);
}
