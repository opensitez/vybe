use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn data_group_two_level_display() {
    let out = run_prints(&p(
        "01 RECORD.\n   05 FIRST-NAME PIC X(5) VALUE \"ALICE\".\n   05 LAST-NAME  PIC X(5) VALUE \"JONES\".",
        "    DISPLAY FIRST-NAME.\n    DISPLAY LAST-NAME.",
    ));
    assert_eq!(out, vec!["ALICE", "JONES"]);
}

#[test]
fn data_group_display_as_whole_group() {
    let out = run_prints(&p(
        "01 PAIR.\n   05 PART-A PIC X(3) VALUE \"ABC\".\n   05 PART-B PIC X(3) VALUE \"DEF\".",
        "    DISPLAY PAIR.",
    ));
    assert_eq!(out, vec!["ABCDEF"]);
}

#[test]
fn data_group_move_to_elementary() {
    let out = run_prints(&p(
        "01 GRP.\n   05 A PIC X(3) VALUE \"XYZ\".\n   05 B PIC X(3) VALUE \"123\".\n01 DST PIC X(6) VALUE SPACES.",
        "    MOVE GRP TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["XYZ123"]);
}

#[test]
fn data_group_three_level_nesting() {
    let out = run_prints(&p(
        "01 LEVEL1.\n   05 LEVEL2.\n      10 LEVEL3 PIC X(4) VALUE \"DEEP\".",
        "    DISPLAY LEVEL3.",
    ));
    assert_eq!(out, vec!["DEEP"]);
}

#[test]
fn data_group_filler_invisible_in_display() {
    let out = run_prints(&p(
        "01 REC.\n   05 FILLER PIC X(3) VALUE \"XXX\".\n   05 DATA-PART PIC X(5) VALUE \"HELLO\".",
        "    DISPLAY DATA-PART.",
    ));
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn data_group_filler_contributes_to_group_size() {
    let out = run_prints(&p(
        "01 REC.\n   05 FILLER PIC X(2) VALUE \"AB\".\n   05 DATA-PART PIC X(3) VALUE \"CDE\".",
        "    DISPLAY REC.",
    ));
    assert_eq!(out, vec!["ABCDE"]);
}

#[test]
fn data_group_modify_elementary_field() {
    let out = run_prints(&p(
        "01 RECORD.\n   05 CODE PIC X(2) VALUE \"AA\".\n   05 VALUE-FIELD PIC 9(4) VALUE 0.",
        "    MOVE \"BB\" TO CODE.\n    MOVE 9999 TO VALUE-FIELD.\n    DISPLAY CODE.\n    DISPLAY VALUE-FIELD.",
    ));
    assert_eq!(out, vec!["BB", "9999"]);
}

#[test]
fn data_group_nested_numeric_and_alpha() {
    let out = run_prints(&p(
        "01 CUSTOMER.\n   05 CUST-ID PIC 9(4) VALUE 1234.\n   05 CUST-NAME PIC X(6) VALUE \"SMITH \".",
        "    DISPLAY CUST-ID.\n    DISPLAY CUST-NAME.",
    ));
    assert_eq!(out, vec!["1234", "SMITH "]);
}

#[test]
fn data_group_multiple_filler_segments() {
    let out = run_prints(&p(
        "01 MSG.\n   05 FILLER PIC X(5) VALUE \"HELLO\".\n   05 FILLER PIC X VALUE \" \".\n   05 FILLER PIC X(5) VALUE \"WORLD\".",
        "    DISPLAY MSG.",
    ));
    assert_eq!(out, vec!["HELLO WORLD"]);
}

#[test]
fn data_group_initialize_resets_all_children() {
    let out = run_prints(&p(
        "01 REC.\n   05 A PIC X(3) VALUE \"ABC\".\n   05 N PIC 9(3) VALUE 999.",
        "    INITIALIZE REC.\n    DISPLAY A.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["   ", "000"]);
}

#[test]
fn data_group_move_into_from_string() {
    let out = run_prints(&p(
        "01 REC.\n   05 A PIC X(3) VALUE \"AAA\".\n   05 B PIC X(3) VALUE \"BBB\".\n01 SRC PIC X(6) VALUE \"XXYYZZ\".",
        "    MOVE SRC TO REC.\n    DISPLAY A.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["XXY", "YZZ"]);
}

#[test]
fn data_group_four_levels_deep() {
    let out = run_prints(&p(
        "01 L1.\n   05 L2.\n      10 L3.\n         15 L4 PIC 9(4) VALUE 4321.",
        "    DISPLAY L4.",
    ));
    assert_eq!(out, vec!["4321"]);
}

#[test]
fn data_group_two_groups_independent() {
    let out = run_prints(&p(
        "01 GRP-A.\n   05 A1 PIC X(3) VALUE \"AAA\".\n   05 A2 PIC X(3) VALUE \"BBB\".\n01 GRP-B.\n   05 B1 PIC X(3) VALUE \"CCC\".\n   05 B2 PIC X(3) VALUE \"DDD\".",
        "    DISPLAY GRP-A.\n    DISPLAY GRP-B.",
    ));
    assert_eq!(out, vec!["AAABBB", "CCCDDD"]);
}

#[test]
fn data_group_compute_child_field() {
    let out = run_prints(&p(
        "01 CALC.\n   05 OPERAND-A PIC 9(4) VALUE 12.\n   05 OPERAND-B PIC 9(4) VALUE 8.\n   05 RESULT PIC 9(6) VALUE 0.",
        "    COMPUTE RESULT = OPERAND-A * OPERAND-B.\n    DISPLAY RESULT.",
    ));
    assert_eq!(out, vec!["96"]);
}

#[test]
fn data_group_with_comp_field() {
    compile_ok(&p(
        "01 REC.\n   05 TEXT PIC X(10) VALUE \"HELLO\".\n   05 BINARY-NUM PIC 9(8) COMP VALUE 0.",
        "    ADD 1 TO BINARY-NUM.",
    ));
}

#[test]
fn data_group_nested_with_occurs() {
    compile_ok(&p(
        "01 OUTER.\n   05 INNER PIC 9(2) OCCURS 5 TIMES.",
        "    MOVE 1 TO INNER(1).\n    MOVE 2 TO INNER(2).",
    ));
}

#[test]
fn data_group_boolean_logic_on_child() {
    let out = run_prints(&p(
        "01 STATUS-REC.\n   05 STATE-CODE PIC X VALUE \"A\".\n   05 STATE-NUM PIC 9 VALUE 1.",
        "    IF STATE-CODE = \"A\" AND STATE-NUM = 1\n        DISPLAY \"VALID\"\n    ELSE\n        DISPLAY \"INVALID\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["VALID"]);
}

#[test]
fn data_group_group_level_condition() {
    let out = run_prints(&p(
        "01 KEY.\n   05 K1 PIC X(2) VALUE \"AB\".\n   05 K2 PIC X(2) VALUE \"CD\".",
        "    IF KEY = \"ABCD\"\n        DISPLAY \"MATCH\"\n    ELSE\n        DISPLAY \"NO MATCH\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["MATCH"]);
}

#[test]
fn data_group_level05_values_display() {
    let out = run_prints(&p(
        "01 EMPLOYEE.\n   05 EMP-NO   PIC 9(6) VALUE 100001.\n   05 EMP-NAME PIC X(10) VALUE \"JOHN      \".\n   05 EMP-DEPT PIC X(3) VALUE \"IT\".",
        "    DISPLAY EMP-NO.\n    DISPLAY EMP-NAME.\n    DISPLAY EMP-DEPT.",
    ));
    assert_eq!(out, vec!["100001", "JOHN      ", "IT "]);
}

#[test]
fn data_group_add_to_nested_numeric() {
    let out = run_prints(&p(
        "01 COUNTERS.\n   05 GOOD-COUNT PIC 9(4) VALUE 100.\n   05 BAD-COUNT PIC 9(4) VALUE 0.",
        "    ADD 1 TO GOOD-COUNT.\n    ADD 1 TO BAD-COUNT.\n    DISPLAY GOOD-COUNT.\n    DISPLAY BAD-COUNT.",
    ));
    assert_eq!(out, vec!["0101", "0001"]);
}

#[test]
fn data_group_copy_between_groups_same_size() {
    let out = run_prints(&p(
        "01 SRC-REC.\n   05 S1 PIC X(3) VALUE \"AAA\".\n   05 S2 PIC X(3) VALUE \"BBB\".\n01 DST-REC.\n   05 D1 PIC X(3) VALUE \"XXX\".\n   05 D2 PIC X(3) VALUE \"YYY\".",
        "    MOVE SRC-REC TO DST-REC.\n    DISPLAY D1.\n    DISPLAY D2.",
    ));
    assert_eq!(out, vec!["AAA", "BBB"]);
}

#[test]
fn data_group_perform_varies_child() {
    let out = run_prints(&p(
        "01 DATA-REC.\n   05 ITER-CNT PIC 9(2) VALUE 0.\n   05 ITER-SUM PIC 9(4) VALUE 0.",
        "    PERFORM VARYING ITER-CNT FROM 1 BY 1 UNTIL ITER-CNT > 5\n        ADD ITER-CNT TO ITER-SUM\n    END-PERFORM.\n    DISPLAY ITER-SUM.",
    ));
    assert_eq!(out, vec!["15"]);
}

#[test]
fn data_group_with_value_low_values() {
    compile_ok(&p(
        "01 REC.\n   05 MARKER PIC X(4) VALUE LOW-VALUES.",
        "    DISPLAY REC.",
    ));
}

#[test]
fn data_group_with_value_high_values() {
    compile_ok(&p(
        "01 REC.\n   05 MARKER PIC X(4) VALUE HIGH-VALUES.",
        "    DISPLAY REC.",
    ));
}

#[test]
fn data_group_level_77_and_01_coexist() {
    let out = run_prints(&p(
        "77 STANDALONE PIC 9(3) VALUE 42.\n01 GRP.\n   05 FIELD PIC 9(3) VALUE 0.",
        "    ADD STANDALONE TO FIELD.\n    DISPLAY FIELD.",
    ));
    assert_eq!(out, vec!["042"]);
}

#[test]
fn data_group_blank_fields_via_filler() {
    let out = run_prints(&p(
        "01 FORMATTED.\n   05 FILLER PIC X(5) VALUE \"ITEM:\".\n   05 FILLER PIC X VALUE \" \".\n   05 ITEM-VALUE PIC 9(3) VALUE 42.",
        "    DISPLAY FORMATTED.",
    ));
    assert_eq!(out, vec!["ITEM:  042"]);
}

#[test]
fn data_group_nested_group_and_elementary() {
    let out = run_prints(&p(
        "01 OUTER.\n   05 HEADER PIC X(4) VALUE \"HEAD\".\n   05 DETAILS.\n      10 D1 PIC 9(2) VALUE 11.\n      10 D2 PIC 9(2) VALUE 22.",
        "    DISPLAY HEADER.\n    DISPLAY D1.\n    DISPLAY D2.",
    ));
    assert_eq!(out, vec!["HEAD", "11", "22"]);
}

#[test]
fn data_group_redisplay_after_move_into_child() {
    let out = run_prints(&p(
        "01 PAIR.\n   05 LEFT PIC X(3) VALUE \"AAA\".\n   05 RIGHT PIC X(3) VALUE \"BBB\".",
        "    MOVE \"ZZZ\" TO LEFT.\n    DISPLAY PAIR.",
    ));
    assert_eq!(out, vec!["ZZZBBB"]);
}
