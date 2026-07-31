use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn compute_multiply_before_add() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = 2 + 3 * 4.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["14"]);
}

#[test]
fn compute_parentheses_override_precedence() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = (2 + 3) * 4.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["20"]);
}

#[test]
fn compute_power_binds_tighter_than_multiply() {
    let out = run_prints(&p(
        "01 R PIC 9(6) VALUE 0.",
        "    COMPUTE R = 2 * 3 ** 2.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["18"]);
}

#[test]
fn compute_double_parentheses() {
    let out = run_prints(&p(
        "01 R PIC 9(6) VALUE 0.",
        "    COMPUTE R = ((4 + 1) * (3 - 1)) ** 2.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["100"]);
}

#[test]
fn compute_chained_subtract() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = 20 - 7 - 3.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["10"]);
}

#[test]
fn compute_mixed_add_subtract() {
    let out = run_prints(&p(
        "01 R PIC S9(5) VALUE 0.",
        "    COMPUTE R = 100 - 40 + 15 - 5.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["70"]);
}

#[test]
fn compute_divide_then_multiply() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = 12 / 4 * 3.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["9"]);
}

#[test]
fn compute_literal_power_of_zero() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = 99 ** 0.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["1"]);
}

#[test]
fn compute_power_of_one() {
    let out = run_prints(&p(
        "01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = 7 ** 1.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["7"]);
}

#[test]
fn compute_left_paren_groups_addition() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = 6 / (2 + 1).\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["2"]);
}

#[test]
fn compute_large_literal_result() {
    let out = run_prints(&p(
        "01 R PIC 9(8) VALUE 0.",
        "    COMPUTE R = 999 * 1000.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["00999000"]);
}

#[test]
fn compute_subtract_to_zero() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = 50 - 50.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["0"]);
}

#[test]
fn compute_using_data_items() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 15.\n01 B PIC 9(3) VALUE 4.\n01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = A * B + A - B.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["71"]);
}

#[test]
fn compute_deeply_nested_parens() {
    let out = run_prints(&p(
        "01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = (((2 + 3) * 4) - 5) * 2.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["30"]);
}

#[test]
fn compute_divide_result_stored_in_decimal_pic() {
    let out = run_prints(&p(
        "01 R PIC 9(3)V99 VALUE 0.",
        "    COMPUTE R = 10 / 4.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["002.50"]);
}

#[test]
fn compute_sum_of_three_vars() {
    let out = run_prints(&p(
        "01 A PIC 9(2) VALUE 10.\n01 B PIC 9(2) VALUE 20.\n01 C PIC 9(2) VALUE 30.\n01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = A + B + C.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["60"]);
}

#[test]
fn compute_subtract_var_from_literal() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 37.\n01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = 100 - A.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["63"]);
}

#[test]
fn compute_multiply_two_vars() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 12.\n01 B PIC 9(3) VALUE 8.\n01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = A * B.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["96"]);
}

#[test]
fn compute_quadratic_expression() {
    // R = 2*x^2 + 3*x + 1 where x=5 => 56
    let out = run_prints(&p(
        "01 X PIC 9(2) VALUE 5.\n01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = 2 * X ** 2 + 3 * X + 1.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["66"]);
}

#[test]
fn compute_integer_square() {
    let out = run_prints(&p(
        "01 N PIC 9(3) VALUE 13.\n01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = N * N.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["169"]);
}

#[test]
fn compute_three_level_nested_divide() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = (100 / (2 + 3)) * 2.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["40"]);
}

#[test]
fn compute_difference_of_products() {
    let out = run_prints(&p(
        "01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = 6 * 7 - 4 * 5.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["22"]);
}

#[test]
fn compute_expression_into_two_receiving_fields() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 0.\n01 B PIC 9(3) VALUE 0.",
        "    COMPUTE A B = 3 + 4.\n    DISPLAY A.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["7", "7"]);
}

#[test]
fn compute_add_decimal_result() {
    let out = run_prints(&p(
        "01 R PIC 9(3)V9 VALUE 0.",
        "    COMPUTE R = 1.5 + 2.5.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["004.0"]);
}

#[test]
fn compute_product_with_decimal_operand() {
    let out = run_prints(&p(
        "01 R PIC 9(4)V9 VALUE 0.",
        "    COMPUTE R = 4 * 2.5.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["0010.0"]);
}

#[test]
fn compute_subtract_decimal_operands() {
    let out = run_prints(&p(
        "01 R PIC 9(3)V9 VALUE 0.",
        "    COMPUTE R = 9.8 - 4.3.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["005.5"]);
}

#[test]
fn compute_with_function_sqrt_rounded() {
    compile_ok(&p(
        "01 R PIC 9(5)V99 VALUE 0.",
        "    COMPUTE R = FUNCTION SQRT(144).",
    ));
}

#[test]
fn compute_power_three() {
    let out = run_prints(&p(
        "01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = 4 ** 3.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["64"]);
}

#[test]
fn compute_subtract_product_from_dividend() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = 50 - 3 * 5.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["35"]);
}

#[test]
fn compute_expression_reused_across_fields() {
    let out = run_prints(&p(
        "01 X PIC 9(2) VALUE 6.\n01 A PIC 9(3) VALUE 0.\n01 B PIC 9(3) VALUE 0.",
        "    COMPUTE A B = X * X - X.\n    DISPLAY A.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["30", "30"]);
}
