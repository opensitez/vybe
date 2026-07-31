use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn condition_and_both_true() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 3.",
        "    IF A > 0 AND B > 0\n        DISPLAY \"BOTH\"\n    ELSE\n        DISPLAY \"NOT BOTH\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["BOTH"]);
}

#[test]
fn condition_and_one_false() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 0.",
        "    IF A > 0 AND B > 0\n        DISPLAY \"BOTH\"\n    ELSE\n        DISPLAY \"NOT BOTH\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["NOT BOTH"]);
}

#[test]
fn condition_or_first_true() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 0.",
        "    IF A > 0 OR B > 0\n        DISPLAY \"EITHER\"\n    ELSE\n        DISPLAY \"NEITHER\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["EITHER"]);
}

#[test]
fn condition_or_both_false() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 0.",
        "    IF A > 0 OR B > 0\n        DISPLAY \"EITHER\"\n    ELSE\n        DISPLAY \"NEITHER\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["NEITHER"]);
}

#[test]
fn condition_not_equal() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 7.",
        "    IF NOT N = 5\n        DISPLAY \"DIFF\"\n    ELSE\n        DISPLAY \"SAME\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["DIFF"]);
}

#[test]
fn condition_not_greater() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 3.",
        "    IF NOT N > 5\n        DISPLAY \"LE\"\n    ELSE\n        DISPLAY \"GT\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["LE"]);
}

#[test]
fn condition_three_way_and() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.\n01 C PIC 9 VALUE 3.",
        "    IF A < B AND B < C AND A < C\n        DISPLAY \"ORDERED\"\n    ELSE\n        DISPLAY \"NOT ORDERED\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ORDERED"]);
}

#[test]
fn condition_and_with_not() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 5.",
        "    IF A > 0 AND NOT B = 0\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn condition_or_with_not() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 0.",
        "    IF NOT A > 0 OR NOT B > 0\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn condition_class_numeric() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"12345\".",
        "    IF S IS NUMERIC\n        DISPLAY \"NUM\"\n    ELSE\n        DISPLAY \"NOT NUM\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["NUM"]);
}

#[test]
fn condition_class_alphabetic() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    IF S IS ALPHABETIC\n        DISPLAY \"ALPHA\"\n    ELSE\n        DISPLAY \"NOT ALPHA\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ALPHA"]);
}

#[test]
fn condition_class_numeric_false_with_letters() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"ABC12\".",
        "    IF S IS NUMERIC\n        DISPLAY \"NUM\"\n    ELSE\n        DISPLAY \"NOT NUM\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["NOT NUM"]);
}

#[test]
fn condition_alphabetic_lower_compiles() {
    compile_ok(&p(
        "01 S PIC X(5) VALUE \"hello\".",
        "    IF S IS ALPHABETIC-LOWER\n        DISPLAY \"LOWER\"\n    END-IF.",
    ));
}

#[test]
fn condition_alphabetic_upper_compiles() {
    compile_ok(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    IF S IS ALPHABETIC-UPPER\n        DISPLAY \"UPPER\"\n    END-IF.",
    ));
}

#[test]
fn condition_abbreviated_relation_or() {
    compile_ok(&p(
        "01 N PIC 9 VALUE 5.",
        "    IF N = 3 OR 5 OR 7\n        DISPLAY \"ODD\"\n    END-IF.",
    ));
}

#[test]
fn condition_abbreviated_relation_and_range() {
    compile_ok(&p(
        "01 N PIC 9(2) VALUE 50.",
        "    IF N >= 10 AND <= 90\n        DISPLAY \"IN RANGE\"\n    END-IF.",
    ));
}

#[test]
fn condition_nested_and_or_precedence() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 0.\n01 C PIC 9 VALUE 1.",
        "    IF A = 1 AND B = 1 OR C = 1\n        DISPLAY \"TRUE\"\n    ELSE\n        DISPLAY \"FALSE\"\n    END-IF.",
    ));
    // AND has higher precedence than OR: (A=1 AND B=1) OR C=1 = FALSE OR TRUE = TRUE
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn condition_not_alphabetic() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"12345\".",
        "    IF NOT S IS ALPHABETIC\n        DISPLAY \"NOT ALPHA\"\n    ELSE\n        DISPLAY \"ALPHA\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["NOT ALPHA"]);
}

#[test]
fn condition_equal_string_comparison() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    IF S = \"HELLO\"\n        DISPLAY \"MATCH\"\n    ELSE\n        DISPLAY \"NO MATCH\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["MATCH"]);
}

#[test]
fn condition_less_than_string() {
    let out = run_prints(&p(
        "01 A PIC X(3) VALUE \"ABC\".\n01 B PIC X(3) VALUE \"DEF\".",
        "    IF A < B\n        DISPLAY \"A BEFORE B\"\n    ELSE\n        DISPLAY \"B BEFORE A\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["A BEFORE B"]);
}

#[test]
fn condition_not_and_combination() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 3.",
        "    IF NOT (X = 1 OR X = 2)\n        DISPLAY \"OTHER\"\n    ELSE\n        DISPLAY \"ONE OR TWO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["OTHER"]);
}

#[test]
fn condition_compound_in_perform() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 0.",
        "    PERFORM UNTIL A >= 3 OR B >= 5\n        ADD 1 TO A\n        ADD 2 TO B\n    END-PERFORM.\n    DISPLAY A.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn condition_pointer_is_not_null_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.",
        "    IF P NOT = NULL\n        DISPLAY \"NOT NULL\"\n    END-IF.",
    ));
}

#[test]
fn condition_double_negation_is_positive() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 5.",
        "    IF NOT NOT (N > 0)\n        DISPLAY \"POS\"\n    ELSE\n        DISPLAY \"NEG\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["POS"]);
}

#[test]
fn condition_not_less_means_ge() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 5.",
        "    IF NOT A < B\n        DISPLAY \"GE\"\n    ELSE\n        DISPLAY \"LT\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["GE"]);
}

#[test]
fn condition_three_or_chain() {
    let out = run_prints(&p(
        "01 C PIC X VALUE \"C\".",
        "    IF C = \"A\" OR C = \"B\" OR C = \"C\"\n        DISPLAY \"ABC\"\n    ELSE\n        DISPLAY \"OTHER\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn condition_numeric_zero_is_not_positive() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.",
        "    IF N > 0\n        DISPLAY \"POS\"\n    ELSE\n        DISPLAY \"ZERO OR NEG\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ZERO OR NEG"]);
}

#[test]
fn condition_complex_boolean_expression() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 0.\n01 C PIC 9 VALUE 5.",
        "    IF (A > 0 AND B = 0) AND C >= 5\n        DISPLAY \"COMPLEX TRUE\"\n    ELSE\n        DISPLAY \"COMPLEX FALSE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["COMPLEX TRUE"]);
}

#[test]
fn condition_or_second_true_only() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 0.\n01 Y PIC 9 VALUE 1.",
        "    IF X > 0 OR Y > 0\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn condition_equal_zero_spaces_numeric() {
    let out = run_prints(&p(
        "01 N PIC 9(3) VALUE ZEROS.",
        "    IF N = ZEROS\n        DISPLAY \"ZERO\"\n    ELSE\n        DISPLAY \"NONZERO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ZERO"]);
}

#[test]
fn condition_equal_spaces_alphanumeric() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE SPACES.",
        "    IF S = SPACES\n        DISPLAY \"BLANK\"\n    ELSE\n        DISPLAY \"NOT BLANK\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["BLANK"]);
}
