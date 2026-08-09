use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn intrinsic_char_from_ord() {
    compile_ok(&p("01 C PIC X.", "    MOVE FUNCTION CHAR(65) TO C."));
}

#[test]
fn intrinsic_ord_of_char() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION ORD(\"A\").\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["65"]);
}

#[test]
fn intrinsic_ord_max_returns_highest() {
    compile_ok(&p(
        "01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = FUNCTION ORD-MAX(\"apple\" \"banana\" \"cherry\").",
    ));
}

#[test]
fn intrinsic_ord_min_returns_lowest() {
    compile_ok(&p(
        "01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = FUNCTION ORD-MIN(\"apple\" \"banana\" \"cherry\").",
    ));
}

#[test]
fn intrinsic_concatenate_two_strings() {
    let out = run_prints(&p(
        "01 R PIC X(15) VALUE SPACES.",
        "    MOVE FUNCTION CONCATENATE(\"HELLO\" \" \" \"WORLD\") TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HELLO WORLD    "]);
}

#[test]
fn intrinsic_concatenate_three() {
    compile_ok(&p(
        "01 R PIC X(20).",
        "    MOVE FUNCTION CONCATENATE(\"A\" \"B\" \"C\") TO R.",
    ));
}

#[test]
fn intrinsic_substitute_single_char() {
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"HELLO\".\n01 R PIC X(10) VALUE SPACES.",
        "    MOVE FUNCTION SUBSTITUTE(S \"L\" \"R\") TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HERRO     "]);
}

#[test]
fn intrinsic_substitute_case_converts() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"Hello\".\n01 R PIC X(10).",
        "    MOVE FUNCTION SUBSTITUTE(S \"H\" \"h\") TO R.",
    ));
}

#[test]
fn intrinsic_trim_trailing_spaces() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"HELLO     \".\n01 R PIC X(10).",
        "    MOVE FUNCTION TRIM(S TRAILING) TO R.",
    ));
}

#[test]
fn intrinsic_trim_both_sides() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"  HELLO  \".\n01 R PIC X(10).",
        "    MOVE FUNCTION TRIM(S) TO R.",
    ));
}

#[test]
fn intrinsic_upper_case_mixed_input() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"Hello\".\n01 R PIC X(5).",
        "    MOVE FUNCTION UPPER-CASE(S) TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn intrinsic_lower_case_mixed_input() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"WORLD\".\n01 R PIC X(5).",
        "    MOVE FUNCTION LOWER-CASE(S) TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["world"]);
}

#[test]
fn intrinsic_reverse_palindrome() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"RADAR\".\n01 R PIC X(5).",
        "    MOVE FUNCTION REVERSE(S) TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["RADAR"]);
}

#[test]
fn intrinsic_length_numeric_field() {
    let out = run_prints(&p(
        "01 N PIC 9(8) VALUE 0.\n01 L PIC 9(4) VALUE 0.",
        "    COMPUTE L = FUNCTION LENGTH(N).\n    DISPLAY L.",
    ));
    assert_eq!(out, vec!["8"]);
}

#[test]
fn intrinsic_current_date_returns_21_chars() {
    compile_ok(&p(
        "01 TODAY PIC X(21).",
        "    MOVE FUNCTION CURRENT-DATE TO TODAY.\n    DISPLAY TODAY.",
    ));
}

#[test]
fn intrinsic_when_compiled_compiles_and_displays() {
    compile_ok(&p(
        "01 COMPILED-AT PIC X(21).",
        "    MOVE FUNCTION WHEN-COMPILED TO COMPILED-AT.\n    DISPLAY COMPILED-AT.",
    ));
}

#[test]
fn intrinsic_char_ord_roundtrip_a() {
    compile_ok(&p(
        "01 C PIC X.\n01 N PIC 9(3) VALUE 0.",
        "    MOVE FUNCTION CHAR(65) TO C.\n    COMPUTE N = FUNCTION ORD(C).",
    ));
}

#[test]
fn intrinsic_concatenate_with_variable() {
    let out = run_prints(&p(
        "01 FIRST PIC X(4) VALUE \"JOHN\".\n01 LAST PIC X(5) VALUE \"SMITH\".\n01 FULL PIC X(10) VALUE SPACES.",
        "    MOVE FUNCTION CONCATENATE(FIRST \" \" LAST) TO FULL.\n    DISPLAY FULL.",
    ));
    assert_eq!(out, vec!["JOHN SMITH"]);
}

#[test]
fn intrinsic_ord_space_is_32() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION ORD(\" \").\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["32"]);
}

#[test]
fn intrinsic_substitute_no_match_unchanged() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".\n01 R PIC X(5) VALUE SPACES.",
        "    MOVE FUNCTION SUBSTITUTE(S \"Z\" \"X\") TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn intrinsic_upper_lower_idempotent() {
    compile_ok(&p(
        "01 S PIC X(5) VALUE \"HELLO\".\n01 R PIC X(5).",
        "    MOVE FUNCTION UPPER-CASE(FUNCTION LOWER-CASE(S)) TO R.",
    ));
}

#[test]
fn intrinsic_reverse_five_chars() {
    let out = run_prints(&p(
        "01 R PIC X(5).",
        "    MOVE FUNCTION REVERSE(\"12345\") TO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["54321"]);
}

#[test]
fn intrinsic_length_empty_string_literal() {
    let out = run_prints(&p(
        "01 L PIC 9(4) VALUE 0.",
        "    COMPUTE L = FUNCTION LENGTH(\"\").\n    DISPLAY L.",
    ));
    assert_eq!(out, vec!["0"]);
}

#[test]
fn intrinsic_concatenate_result_length() {
    let out = run_prints(&p(
        "01 R PIC X(10).\n01 L PIC 9(4) VALUE 0.",
        "    MOVE FUNCTION CONCATENATE(\"ABC\" \"DE\") TO R.\n    COMPUTE L = FUNCTION LENGTH(FUNCTION CONCATENATE(\"ABC\" \"DE\")).\n    DISPLAY L.",
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn intrinsic_substitute_multi_char_old() {
    compile_ok(&p(
        "01 S PIC X(20) VALUE \"HELLO WORLD\".\n01 R PIC X(20).",
        "    MOVE FUNCTION SUBSTITUTE(S \"WORLD\" \"COBOL\") TO R.",
    ));
}

#[test]
fn intrinsic_formatted_date_compiles() {
    compile_ok(&p(
        "01 D PIC X(10).",
        "    MOVE FUNCTION FORMATTED-DATE(\"YYYY-MM-DD\" 20230615) TO D.",
    ));
}

#[test]
fn intrinsic_formatted_datetime_compiles() {
    compile_ok(&p(
        "01 DT PIC X(30).",
        "    MOVE FUNCTION FORMATTED-DATETIME(\"YYYY-MM-DDThh:mm:ss\" 20230615 1430) TO DT.",
    ));
}

#[test]
fn intrinsic_numval_simple_integer() {
    compile_ok(&p(
        "01 S PIC X(5) VALUE \"42\".\n01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = FUNCTION NUMVAL(S).",
    ));
}

#[test]
fn intrinsic_combined_upper_and_reverse() {
    compile_ok(&p(
        "01 S PIC X(5) VALUE \"hello\".\n01 R PIC X(5).",
        "    MOVE FUNCTION REVERSE(FUNCTION UPPER-CASE(S)) TO R.",
    ));
}
