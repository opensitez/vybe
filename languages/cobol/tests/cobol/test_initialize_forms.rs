use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn initialize_alphanumeric_to_spaces() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    INITIALIZE S.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["     "]);
}

#[test]
fn initialize_numeric_to_zero() {
    let out = run_prints(&p(
        "01 N PIC 9(5) VALUE 99999.",
        "    INITIALIZE N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["00000"]);
}

#[test]
fn initialize_group_item_all_fields() {
    let out = run_prints(&p(
        "01 GRP.\n   05 A PIC X(3) VALUE \"ABC\".\n   05 B PIC 9(3) VALUE 123.",
        "    INITIALIZE GRP.\n    DISPLAY A.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["   ", "000"]);
}

#[test]
fn initialize_multiple_fields() {
    let out = run_prints(&p(
        "01 A PIC X(4) VALUE \"XXXX\".\n01 B PIC 9(4) VALUE 1111.",
        "    INITIALIZE A B.\n    DISPLAY A.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["    ", "0000"]);
}

#[test]
fn initialize_alphabetic_category() {
    compile_ok(&p(
        "01 S PIC A(5) VALUE \"HELLO\".",
        "    INITIALIZE S REPLACING ALPHABETIC DATA BY \"X\".",
    ));
}

#[test]
fn initialize_numeric_category() {
    compile_ok(&p(
        "01 N PIC 9(5) VALUE 12345.",
        "    INITIALIZE N REPLACING NUMERIC DATA BY 9.",
    ));
}

#[test]
fn initialize_alphanumeric_category_replacing() {
    compile_ok(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    INITIALIZE S REPLACING ALPHANUMERIC DATA BY \"_\".",
    ));
}

#[test]
fn initialize_signed_numeric() {
    let out = run_prints(&p(
        "01 N PIC S9(4) VALUE -999.",
        "    INITIALIZE N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["+0000"]);
}

#[test]
fn initialize_decimal_pic() {
    let out = run_prints(&p(
        "01 D PIC 9(3)V99 VALUE 123.45.",
        "    INITIALIZE D.\n    DISPLAY D.",
    ));
    assert_eq!(out, vec!["00000"]);
}

#[test]
fn initialize_then_use() {
    let out = run_prints(&p(
        "01 N PIC 9(4) VALUE 9999.",
        "    INITIALIZE N.\n    ADD 42 TO N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0042"]);
}

#[test]
fn initialize_table_element() {
    compile_ok(&p(
        "01 T.\n   05 E PIC 9(3) OCCURS 5 TIMES.",
        "    INITIALIZE T.",
    ));
}

#[test]
fn initialize_level77_field() {
    let out = run_prints(&p(
        "77 WS-VAL PIC 9(4) VALUE 5000.",
        "    INITIALIZE WS-VAL.\n    DISPLAY WS-VAL.",
    ));
    assert_eq!(out, vec!["0000"]);
}

#[test]
fn initialize_before_loop_counter() {
    let out = run_prints(&p(
        "01 C PIC 9(3) VALUE 999.",
        "    INITIALIZE C.\n    PERFORM UNTIL C >= 5\n        ADD 1 TO C\n    END-PERFORM.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["005"]);
}

#[test]
fn initialize_resets_between_operations() {
    let out = run_prints(&p(
        "01 S PIC 9(3) VALUE 0.",
        "    ADD 50 TO S.\n    INITIALIZE S.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["000"]);
}

#[test]
fn initialize_group_nested_levels() {
    let out = run_prints(&p(
        "01 OUTER.\n   05 INNER.\n      10 DEEPEST PIC 9(2) VALUE 99.",
        "    INITIALIZE OUTER.\n    DISPLAY DEEPEST.",
    ));
    assert_eq!(out, vec!["00"]);
}

#[test]
fn set_index_from_integer() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X OCCURS 10 TIMES INDEXED BY IDX.",
        "    SET IDX TO 1.",
    ));
}

#[test]
fn set_index_up_by() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X OCCURS 10 TIMES INDEXED BY IDX.",
        "    SET IDX TO 1.\n    SET IDX UP BY 3.",
    ));
}

#[test]
fn set_index_down_by() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X OCCURS 10 TIMES INDEXED BY IDX.",
        "    SET IDX TO 5.\n    SET IDX DOWN BY 2.",
    ));
}

#[test]
fn set_boolean_flag_to_true() {
    let out = run_prints(&p(
        "01 F PIC X VALUE \"N\".\n    88 F-YES VALUE \"Y\".",
        "    SET F-YES TO TRUE.\n    DISPLAY F.",
    ));
    assert_eq!(out, vec!["Y"]);
}

#[test]
fn set_condition_name_then_test() {
    let out = run_prints(&p(
        "01 MODE PIC X(4) VALUE \"IDLE\".\n    88 RUNNING VALUE \"WORK\".\n    88 IDLE VALUE \"IDLE\".",
        "    SET RUNNING TO TRUE.\n    IF RUNNING\n        DISPLAY \"RUNNING\"\n    ELSE\n        DISPLAY \"IDLE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["RUNNING"]);
}

#[test]
fn set_pointer_to_null() {
    compile_ok(&p(
        "01 PTR USAGE POINTER.",
        "    SET PTR TO NULL.",
    ));
}

#[test]
fn set_condition_false_and_verify() {
    let out = run_prints(&p(
        "01 F PIC X VALUE \"Y\".\n    88 F-ON VALUE \"Y\".\n    88 F-OFF VALUE \"N\".",
        "    SET F-ON TO FALSE.\n    IF F-OFF\n        DISPLAY \"OFF\"\n    ELSE\n        DISPLAY \"ON\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["OFF"]);
}

#[test]
fn set_multiple_conditions_same_field() {
    let out = run_prints(&p(
        "01 STATUS PIC 9 VALUE 0.\n    88 OPEN-STATUS VALUE 1.\n    88 CLOSE-STATUS VALUE 2.",
        "    SET OPEN-STATUS TO TRUE.\n    DISPLAY STATUS.",
    ));
    assert_eq!(out, vec!["1"]);
}

#[test]
fn initialize_and_set_in_sequence() {
    let out = run_prints(&p(
        "01 N PIC 9(4) VALUE 9999.\n01 FLAG PIC X VALUE \"N\".\n    88 DONE VALUE \"Y\".",
        "    INITIALIZE N.\n    SET DONE TO TRUE.\n    IF DONE\n        DISPLAY N\n    END-IF.",
    ));
    assert_eq!(out, vec!["0000"]);
}

#[test]
fn initialize_replacing_numeric_by_literal() {
    compile_ok(&p(
        "01 A PIC 9(3).\n01 B PIC 9(3).",
        "    INITIALIZE A B REPLACING NUMERIC DATA BY 5.",
    ));
}

#[test]
fn set_condition_inside_if() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 1.\n01 FLAG PIC X VALUE \"N\".\n    88 FOUND VALUE \"Y\".",
        "    IF N > 0\n        SET FOUND TO TRUE\n    END-IF.\n    IF FOUND\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn initialize_resets_multiple_times() {
    let out = run_prints(&p(
        "01 N PIC 9(3) VALUE 0.",
        "    ADD 100 TO N.\n    INITIALIZE N.\n    ADD 50 TO N.\n    INITIALIZE N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["000"]);
}

#[test]
fn set_index_to_length_of_table() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X OCCURS 5 TIMES INDEXED BY IX.",
        "    SET IX TO 5.",
    ));
}

#[test]
fn set_multiple_indexes_same_value() {
    compile_ok(&p(
        "01 T1.\n   05 E1 PIC X OCCURS 5 TIMES INDEXED BY IX1.\n01 T2.\n   05 E2 PIC X OCCURS 5 TIMES INDEXED BY IX2.",
        "    SET IX1 IX2 TO 1.",
    ));
}

#[test]
fn initialize_all_working_storage_group() {
    let out = run_prints(&p(
        "01 REC.\n   05 NAME PIC X(5) VALUE \"HELLO\".\n   05 AGE PIC 9(3) VALUE 42.\n   05 CODE PIC X(2) VALUE \"AB\".",
        "    INITIALIZE REC.\n    DISPLAY NAME.\n    DISPLAY AGE.\n    DISPLAY CODE.",
    ));
    assert_eq!(out, vec!["     ", "000", "  "]);
}
