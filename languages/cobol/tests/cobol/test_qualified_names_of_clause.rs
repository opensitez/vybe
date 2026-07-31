use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn qualified_of_clause_basic() {
    let out = run_prints(&p(
        "01 GRP-A.\n   05 CODE PIC X(3) VALUE \"AAA\".\n01 GRP-B.\n   05 CODE PIC X(3) VALUE \"BBB\".",
        "    DISPLAY CODE OF GRP-A.\n    DISPLAY CODE OF GRP-B.",
    ));
    assert_eq!(out, vec!["AAA", "BBB"]);
}

#[test]
fn qualified_of_move_into_qualified() {
    let out = run_prints(&p(
        "01 REC-1.\n   05 VALUE-FIELD PIC 9(4) VALUE 1111.\n01 REC-2.\n   05 VALUE-FIELD PIC 9(4) VALUE 2222.",
        "    MOVE VALUE-FIELD OF REC-1 TO VALUE-FIELD OF REC-2.\n    DISPLAY VALUE-FIELD OF REC-2.",
    ));
    assert_eq!(out, vec!["1111"]);
}

#[test]
fn qualified_of_add_to_qualified() {
    let out = run_prints(&p(
        "01 ACC-A.\n   05 TOTAL PIC 9(5) VALUE 100.\n01 ACC-B.\n   05 TOTAL PIC 9(5) VALUE 200.",
        "    ADD TOTAL OF ACC-A TO TOTAL OF ACC-B.\n    DISPLAY TOTAL OF ACC-B.",
    ));
    assert_eq!(out, vec!["00300"]);
}

#[test]
fn qualified_three_level_deep() {
    let out = run_prints(&p(
        "01 OUTER.\n   05 MIDDLE.\n      10 INNER PIC X(5) VALUE \"DEEP\".",
        "    DISPLAY INNER OF MIDDLE OF OUTER.",
    ));
    assert_eq!(out, vec!["DEEP "]);
}

#[test]
fn qualified_two_records_same_subfield() {
    let out = run_prints(&p(
        "01 EMP-A.\n   05 NAME PIC X(8) VALUE \"ALICE   \".\n   05 DEPT PIC X(3) VALUE \"IT\".\n01 EMP-B.\n   05 NAME PIC X(8) VALUE \"BOB     \".\n   05 DEPT PIC X(3) VALUE \"HR\".",
        "    DISPLAY NAME OF EMP-A.\n    DISPLAY DEPT OF EMP-B.",
    ));
    assert_eq!(out, vec!["ALICE   ", "HR "]);
}

#[test]
fn qualified_compare_two_same_named_fields() {
    let out = run_prints(&p(
        "01 OLD-REC.\n   05 STATUS PIC X VALUE \"A\".\n01 NEW-REC.\n   05 STATUS PIC X VALUE \"B\".",
        "    IF STATUS OF OLD-REC NOT = STATUS OF NEW-REC\n        DISPLAY \"CHANGED\"\n    ELSE\n        DISPLAY \"SAME\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["CHANGED"]);
}

#[test]
fn qualified_in_evaluate_subject() {
    let out = run_prints(&p(
        "01 REC-X.\n   05 TYPE PIC X VALUE \"A\".\n01 REC-Y.\n   05 TYPE PIC X VALUE \"B\".",
        "    EVALUATE TYPE OF REC-X\n        WHEN \"A\"\n            DISPLAY \"TYPE A\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["TYPE A"]);
}

#[test]
fn qualified_in_string_concat() {
    compile_ok(&p(
        "01 SRC-A.\n   05 WORD PIC X(5) VALUE \"HELLO\".\n01 SRC-B.\n   05 WORD PIC X(5) VALUE \"WORLD\".\n01 R PIC X(15) VALUE SPACES.",
        "    STRING WORD OF SRC-A DELIMITED BY SPACE \" \" DELIMITED BY SIZE WORD OF SRC-B DELIMITED BY SPACE INTO R.",
    ));
}

#[test]
fn qualified_of_inside_loop() {
    let out = run_prints(&p(
        "01 COUNTERS.\n   05 CNT PIC 9(3) VALUE 0.\n01 TOTALS.\n   05 CNT PIC 9(3) VALUE 0.",
        "    PERFORM 3 TIMES\n        ADD 1 TO CNT OF COUNTERS\n    END-PERFORM.\n    DISPLAY CNT OF COUNTERS.",
    ));
    assert_eq!(out, vec!["003"]);
}

#[test]
fn qualified_two_levels_in_condition() {
    let out = run_prints(&p(
        "01 ORDER-HDR.\n   05 ORDER-NO PIC 9(5) VALUE 10001.\n01 LINE-ITEM.\n   05 ORDER-NO PIC 9(5) VALUE 10001.",
        "    IF ORDER-NO OF ORDER-HDR = ORDER-NO OF LINE-ITEM\n        DISPLAY \"MATCH\"\n    ELSE\n        DISPLAY \"MISMATCH\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["MATCH"]);
}

#[test]
fn qualified_field_in_compute() {
    let out = run_prints(&p(
        "01 BASE-DATA.\n   05 PRICE PIC 9(5)V99 VALUE 10.50.\n01 DERIVED.\n   05 PRICE PIC 9(5)V99 VALUE 0.\n01 R PIC 9(6)V99 VALUE 0.",
        "    COMPUTE R = PRICE OF BASE-DATA * 2.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["21.00"]);
}

#[test]
fn qualified_three_ambiguous_names() {
    let out = run_prints(&p(
        "01 GRP1.\n   05 X PIC 9 VALUE 1.\n01 GRP2.\n   05 X PIC 9 VALUE 2.\n01 GRP3.\n   05 X PIC 9 VALUE 3.",
        "    DISPLAY X OF GRP1.\n    DISPLAY X OF GRP2.\n    DISPLAY X OF GRP3.",
    ));
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn qualified_move_chained() {
    let out = run_prints(&p(
        "01 A.\n   05 VAL PIC 9(3) VALUE 100.\n01 B.\n   05 VAL PIC 9(3) VALUE 0.\n01 C.\n   05 VAL PIC 9(3) VALUE 0.",
        "    MOVE VAL OF A TO VAL OF B.\n    MOVE VAL OF B TO VAL OF C.\n    DISPLAY VAL OF C.",
    ));
    assert_eq!(out, vec!["100"]);
}

#[test]
fn qualified_in_perform_until() {
    let out = run_prints(&p(
        "01 LOOP-CTRL.\n   05 LIMIT PIC 9(2) VALUE 5.\n01 STATE.\n   05 LIMIT PIC 9(2) VALUE 0.",
        "    PERFORM UNTIL LIMIT OF STATE >= LIMIT OF LOOP-CTRL\n        ADD 1 TO LIMIT OF STATE\n    END-PERFORM.\n    DISPLAY LIMIT OF STATE.",
    ));
    assert_eq!(out, vec!["05"]);
}

#[test]
fn qualified_nested_four_levels() {
    compile_ok(&p(
        "01 L1.\n   05 L2.\n      10 L3.\n         15 NAME PIC X(5) VALUE \"COBOL\".",
        "    DISPLAY NAME OF L3 OF L2 OF L1.",
    ));
}

#[test]
fn qualified_in_inspect() {
    compile_ok(&p(
        "01 STR-A.\n   05 DATA PIC X(10) VALUE \"HELLO\".\n01 STR-B.\n   05 DATA PIC X(10) VALUE \"WORLD\".\n01 CNT PIC 9(2) VALUE 0.",
        "    INSPECT DATA OF STR-A TALLYING CNT FOR ALL \"L\".",
    ));
}

#[test]
fn qualified_add_multiple_qualified_sources() {
    let out = run_prints(&p(
        "01 BUDGET-A.\n   05 AMOUNT PIC 9(6) VALUE 1000.\n01 BUDGET-B.\n   05 AMOUNT PIC 9(6) VALUE 2000.\n01 TOTAL PIC 9(7) VALUE 0.",
        "    ADD AMOUNT OF BUDGET-A AMOUNT OF BUDGET-B GIVING TOTAL.\n    DISPLAY TOTAL.",
    ));
    assert_eq!(out, vec!["3000"]);
}

#[test]
fn qualified_initialize_one_of_two_same_fields() {
    let out = run_prints(&p(
        "01 REC-1.\n   05 SCORE PIC 9(3) VALUE 100.\n01 REC-2.\n   05 SCORE PIC 9(3) VALUE 200.",
        "    INITIALIZE SCORE OF REC-1.\n    DISPLAY SCORE OF REC-1.\n    DISPLAY SCORE OF REC-2.",
    ));
    assert_eq!(out, vec!["000", "200"]);
}

#[test]
fn qualified_field_display_then_move_back() {
    let out = run_prints(&p(
        "01 SOURCE.\n   05 ID PIC 9(4) VALUE 9999.\n01 TARGET.\n   05 ID PIC 9(4) VALUE 0.",
        "    DISPLAY ID OF SOURCE.\n    MOVE ID OF SOURCE TO ID OF TARGET.\n    DISPLAY ID OF TARGET.",
    ));
    assert_eq!(out, vec!["9999", "9999"]);
}

#[test]
fn qualified_string_delimited_field() {
    compile_ok(&p(
        "01 PART-A.\n   05 LABEL PIC X(10) VALUE \"HELLO     \".\n01 PART-B.\n   05 LABEL PIC X(10) VALUE \"WORLD     \".\n01 RESULT PIC X(25) VALUE SPACES.",
        "    STRING LABEL OF PART-A DELIMITED BY SPACE \" \" DELIMITED BY SIZE LABEL OF PART-B DELIMITED BY SPACE INTO RESULT.",
    ));
}

#[test]
fn qualified_two_fields_evaluated_together() {
    let out = run_prints(&p(
        "01 HEADER.\n   05 TYPE-CODE PIC X VALUE \"A\".\n01 DETAIL.\n   05 TYPE-CODE PIC X VALUE \"D\".",
        "    EVALUATE TYPE-CODE OF HEADER ALSO TYPE-CODE OF DETAIL\n        WHEN \"A\" ALSO \"D\"\n            DISPLAY \"VALID\"\n        WHEN OTHER ALSO OTHER\n            DISPLAY \"INVALID\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["VALID"]);
}

#[test]
fn qualified_level88_in_qualified_field() {
    let out = run_prints(&p(
        "01 FLAG-REC.\n   05 ACTIVE-FLAG PIC X VALUE \"Y\".\n      88 ACTIVE-ON VALUE \"Y\".",
        "    IF ACTIVE-ON\n        DISPLAY \"ACTIVE\"\n    ELSE\n        DISPLAY \"INACTIVE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ACTIVE"]);
}

#[test]
fn qualified_add_to_same_named_fields_separately() {
    let out = run_prints(&p(
        "01 DEPT-A.\n   05 HEADCOUNT PIC 9(3) VALUE 5.\n01 DEPT-B.\n   05 HEADCOUNT PIC 9(3) VALUE 10.",
        "    ADD 1 TO HEADCOUNT OF DEPT-A.\n    ADD 2 TO HEADCOUNT OF DEPT-B.\n    DISPLAY HEADCOUNT OF DEPT-A.\n    DISPLAY HEADCOUNT OF DEPT-B.",
    ));
    assert_eq!(out, vec!["006", "012"]);
}

#[test]
fn qualified_different_pics_same_name() {
    let out = run_prints(&p(
        "01 REC-ALPHA.\n   05 KEY PIC X(4) VALUE \"ABCD\".\n01 REC-NUM.\n   05 KEY PIC 9(4) VALUE 1234.",
        "    DISPLAY KEY OF REC-ALPHA.\n    DISPLAY KEY OF REC-NUM.",
    ));
    assert_eq!(out, vec!["ABCD", "1234"]);
}

#[test]
fn qualified_inspect_two_fields() {
    compile_ok(&p(
        "01 TEXT-A.\n   05 CONTENT PIC X(10) VALUE \"HELLO\".\n01 TEXT-B.\n   05 CONTENT PIC X(10) VALUE \"WORLD\".\n01 CNT-A PIC 9(2) VALUE 0.\n01 CNT-B PIC 9(2) VALUE 0.",
        "    INSPECT CONTENT OF TEXT-A TALLYING CNT-A FOR ALL \"L\".\n    INSPECT CONTENT OF TEXT-B TALLYING CNT-B FOR ALL \"O\".",
    ));
}

#[test]
fn qualified_compute_using_two_qualified() {
    let out = run_prints(&p(
        "01 SALARY-TABLE.\n   05 BASE-PAY PIC 9(6) VALUE 50000.\n01 BONUS-TABLE.\n   05 BASE-PAY PIC 9(6) VALUE 10000.\n01 TOTAL-COMP PIC 9(7) VALUE 0.",
        "    COMPUTE TOTAL-COMP = BASE-PAY OF SALARY-TABLE + BASE-PAY OF BONUS-TABLE.\n    DISPLAY TOTAL-COMP.",
    ));
    assert_eq!(out, vec!["60000"]);
}

#[test]
fn qualified_subtract_from_two_targets() {
    let out = run_prints(&p(
        "01 GROUP-X.\n   05 BALANCE PIC 9(5) VALUE 1000.\n01 GROUP-Y.\n   05 BALANCE PIC 9(5) VALUE 2000.",
        "    SUBTRACT 100 FROM BALANCE OF GROUP-X.\n    SUBTRACT 500 FROM BALANCE OF GROUP-Y.\n    DISPLAY BALANCE OF GROUP-X.\n    DISPLAY BALANCE OF GROUP-Y.",
    ));
    assert_eq!(out, vec!["00900", "01500"]);
}
