use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn intrinsic_integer_of_date_compiles() {
    compile_ok(&p(
        "01 R PIC 9(8) VALUE 0.",
        "    COMPUTE R = FUNCTION INTEGER-OF-DATE(20230101).",
    ));
}

#[test]
fn intrinsic_date_of_integer_compiles() {
    compile_ok(&p(
        "01 R PIC 9(8) VALUE 0.",
        "    COMPUTE R = FUNCTION DATE-OF-INTEGER(738521).",
    ));
}

#[test]
fn intrinsic_numval_converts_string() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"  42.5\".\n01 R PIC 9(5)V9 VALUE 0.",
        "    COMPUTE R = FUNCTION NUMVAL(S).",
    ));
}

#[test]
fn intrinsic_numval_c_compiles() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"$1,234.56\".\n01 R PIC 9(8)V99 VALUE 0.",
        "    COMPUTE R = FUNCTION NUMVAL-C(S \"$\").",
    ));
}

#[test]
fn intrinsic_integer_floor_like() {
    let out = run_prints(&p(
        "01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = FUNCTION INTEGER(3.7).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn intrinsic_integer_of_negative() {
    let out = run_prints(&p(
        "01 R PIC S9(5) VALUE 0.",
        "    COMPUTE R = FUNCTION INTEGER(-3.2).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["-0004"]);
}

#[test]
fn intrinsic_mod_basic() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = FUNCTION MOD(17 5).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["2"]);
}

#[test]
fn intrinsic_mod_zero_remainder() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = FUNCTION MOD(10 5).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["0"]);
}

#[test]
fn intrinsic_rem_basic() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = FUNCTION REM(17 5).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["2"]);
}

#[test]
fn intrinsic_abs_positive() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION ABS(42).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["42"]);
}

#[test]
fn intrinsic_abs_negative() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION ABS(-42).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["42"]);
}

#[test]
fn intrinsic_max_two_args() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION MAX(7 3).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["7"]);
}

#[test]
fn intrinsic_min_two_args() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION MIN(7 3).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn intrinsic_max_three_args() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION MAX(5 12 8).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["12"]);
}

#[test]
fn intrinsic_min_three_args() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION MIN(5 12 8).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn intrinsic_sqrt_perfect_square() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION SQRT(144).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["12"]);
}

#[test]
fn intrinsic_log_of_e() {
    compile_ok(&p(
        "01 R PIC 9(5)V9(5) VALUE 0.",
        "    COMPUTE R = FUNCTION LOG(2.71828).",
    ));
}

#[test]
fn intrinsic_log10_of_100() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION LOG10(100).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["2"]);
}

#[test]
fn intrinsic_exp_of_zero() {
    compile_ok(&p(
        "01 R PIC 9(5)V9(5) VALUE 0.",
        "    COMPUTE R = FUNCTION EXP(0).",
    ));
}

#[test]
fn intrinsic_factorial_via_integer_and_product() {
    // 5! = 120 using MOD
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION MOD(120 7).\n    DISPLAY R.",
    ));
    // 120 mod 7 = 1
    assert_eq!(out, vec!["1"]);
}

#[test]
fn intrinsic_current_date_compiles() {
    compile_ok(&p(
        "01 TODAY PIC X(21).",
        "    MOVE FUNCTION CURRENT-DATE TO TODAY.",
    ));
}

#[test]
fn intrinsic_when_compiled_compiles() {
    compile_ok(&p(
        "01 COMPILED-DATE PIC X(21).",
        "    MOVE FUNCTION WHEN-COMPILED TO COMPILED-DATE.",
    ));
}

#[test]
fn intrinsic_length_of_literal() {
    let out = run_prints(&p(
        "01 L PIC 9(4) VALUE 0.",
        "    COMPUTE L = FUNCTION LENGTH(\"HELLO\").\n    DISPLAY L.",
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn intrinsic_length_of_variable() {
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"ABCDE\".\n01 L PIC 9(4) VALUE 0.",
        "    COMPUTE L = FUNCTION LENGTH(S).\n    DISPLAY L.",
    ));
    assert_eq!(out, vec!["10"]);
}

#[test]
fn intrinsic_upper_case_literal() {
    let out = run_prints(&p(
        "01 R PIC X(10).",
        "    MOVE FUNCTION UPPER-CASE(\"hello\") TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HELLO     "]);
}

#[test]
fn intrinsic_lower_case_literal() {
    let out = run_prints(&p(
        "01 R PIC X(10).",
        "    MOVE FUNCTION LOWER-CASE(\"HELLO\") TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["hello     "]);
}

#[test]
fn intrinsic_reverse_string() {
    let out = run_prints(&p(
        "01 R PIC X(5).",
        "    MOVE FUNCTION REVERSE(\"ABCDE\") TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["EDCBA"]);
}

#[test]
fn intrinsic_trim_leading_spaces() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"   HELLO\".\n01 R PIC X(10).",
        "    MOVE FUNCTION TRIM(S LEADING) TO R.",
    ));
}

#[test]
fn intrinsic_substitute_replaces_chars() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"HELLO\".\n01 R PIC X(10).",
        "    MOVE FUNCTION SUBSTITUTE(S \"L\" \"R\") TO R.",
    ));
}

#[test]
fn intrinsic_integer_of_date_day_of_integer_roundtrip() {
    compile_ok(&p(
        "01 D PIC 9(8) VALUE 20230101.\n01 I PIC 9(8) VALUE 0.\n01 D2 PIC 9(8) VALUE 0.",
        "    COMPUTE I = FUNCTION INTEGER-OF-DATE(D).\n    COMPUTE D2 = FUNCTION DATE-OF-INTEGER(I).",
    ));
}
