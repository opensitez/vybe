use crate::helpers::run_print;

#[test]
fn complex_literal_imaginary_unit() {
    assert_eq!(run_print("1j"), "1j");
}

#[test]
fn complex_literal_with_real_part() {
    assert_eq!(run_print("3+4j"), "(3+4j)");
}

#[test]
fn complex_literal_negative_imag() {
    assert_eq!(run_print("2-3j"), "(2-3j)");
}

#[test]
fn complex_addition() {
    assert_eq!(run_print("(1+2j) + (3+4j)"), "(4+6j)");
}

#[test]
fn complex_subtraction() {
    assert_eq!(run_print("(5+3j) - (2+1j)"), "(3+2j)");
}

#[test]
fn complex_multiplication() {
    assert_eq!(run_print("(1+2j) * (3+4j)"), "(-5+10j)");
}

#[test]
fn complex_real_part_access() {
    assert_eq!(run_print("(3+4j).real"), "3.0");
}

#[test]
fn complex_imag_part_access() {
    assert_eq!(run_print("(3+4j).imag"), "4.0");
}

#[test]
fn complex_conjugate() {
    assert_eq!(run_print("(3+4j).conjugate()"), "(3-4j)");
}

#[test]
fn complex_abs_magnitude() {
    assert_eq!(run_print("abs(3+4j)"), "5.0");
}

#[test]
fn complex_equality() {
    assert_eq!(run_print("(1+2j) == (1+2j)"), "True");
}

#[test]
fn complex_inequality() {
    assert_eq!(run_print("(1+2j) != (2+1j)"), "True");
}

#[test]
fn complex_negation() {
    assert_eq!(run_print("-(1+1j)"), "(-1-1j)");
}

#[test]
fn complex_multiply_by_real() {
    assert_eq!(run_print("(2+3j) * 2"), "(4+6j)");
}

#[test]
fn complex_division_by_real() {
    assert_eq!(run_print("(4+6j) / 2"), "(2+3j)");
}

#[test]
fn complex_pure_imaginary_squared() {
    assert_eq!(run_print("(1j) ** 2"), "(-1+0j)");
}

#[test]
fn complex_zero() {
    assert_eq!(run_print("0+0j"), "0j");
}

#[test]
fn complex_bool_nonzero_true() {
    assert_eq!(run_print("bool(1+1j)"), "True");
}

#[test]
fn complex_bool_zero_false() {
    assert_eq!(run_print("bool(0j)"), "False");
}

#[test]
fn complex_in_list() {
    assert_eq!(run_print("[1j, 2+3j]"), "[1j, (2+3j)]");
}

#[test]
fn complex_float_real_part() {
    assert_eq!(run_print("(2.5+0j).real"), "2.5");
}

#[test]
fn complex_add_real_int_promotes() {
    assert_eq!(run_print("(1+2j) + 3"), "(4+2j)");
}

#[test]
fn complex_sub_real() {
    assert_eq!(run_print("(5+2j) - 1"), "(4+2j)");
}

#[test]
fn complex_repr_has_j_suffix() {
    assert_eq!(run_print("repr(1+0j)"), "(1+0j)");
}

#[test]
fn complex_literal_uppercase_j() {
    assert_eq!(run_print("2+3J"), "(2+3j)");
}

#[test]
fn complex_multiplication_by_j_rotates() {
    assert_eq!(run_print("1j * 1j"), "(-1+0j)");
}

#[test]
fn complex_equality_real_parts_only_differs() {
    assert_eq!(run_print("(1+0j) == 1"), "False");
}

#[test]
fn complex_sum_of_conjugates_is_real() {
    assert_eq!(run_print("(3+4j) + (3-4j)"), "(6+0j)");
}

#[test]
fn complex_imag_part_after_addition() {
    assert_eq!(run_print("((1+2j) + (3-2j)).imag"), "0.0");
}

#[test]
fn complex_nested_expression() {
    assert_eq!(run_print("((1+1j) * 2) + 1"), "(3+2j)");
}
