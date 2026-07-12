use crate::helpers::{run_print, run_python_one};

#[test]
fn floordiv_positive_by_positive() {
    assert_eq!(run_print("7 // 2"), "3");
}

#[test]
fn floordiv_negative_dividend() {
    assert_eq!(run_print("-7 // 2"), "-4");
}

#[test]
fn floordiv_negative_divisor() {
    assert_eq!(run_print("7 // -2"), "-4");
}

#[test]
fn floordiv_both_negative() {
    assert_eq!(run_print("-7 // -2"), "3");
}

#[test]
fn floordiv_exact_division() {
    assert_eq!(run_print("10 // 5"), "2");
}

#[test]
fn floordiv_by_one() {
    assert_eq!(run_print("42 // 1"), "42");
}

#[test]
fn floordiv_zero_dividend() {
    assert_eq!(run_print("0 // 5"), "0");
}

#[test]
fn modulo_positive_remainder() {
    assert_eq!(run_print("7 % 3"), "1");
}

#[test]
fn modulo_exact_multiple() {
    assert_eq!(run_print("10 % 5"), "0");
}

#[test]
fn modulo_negative_dividend() {
    assert_eq!(run_print("-7 % 3"), "2");
}

#[test]
fn modulo_negative_divisor() {
    assert_eq!(run_print("7 % -3"), "-2");
}

#[test]
fn divmod_returns_quotient_and_remainder() {
    assert_eq!(run_print("divmod(7, 3)"), "(2, 1)");
}

#[test]
fn divmod_exact_multiple() {
    assert_eq!(run_print("divmod(10, 5)"), "(2, 0)");
}

#[test]
fn divmod_negative_dividend() {
    assert_eq!(run_print("divmod(-7, 3)"), "(-3, 2)");
}

#[test]
fn divmod_negative_divisor() {
    assert_eq!(run_print("divmod(7, -3)"), "(-3, -2)");
}

#[test]
fn divmod_both_negative() {
    assert_eq!(run_print("divmod(-7, -3)"), "(2, -1)");
}

#[test]
fn divmod_large_numbers() {
    assert_eq!(run_print("divmod(100, 7)"), "(14, 2)");
}

#[test]
fn floordiv_float_truncates_toward_negative_infinity() {
    assert_eq!(run_print("7.5 // 2"), "3.0");
}

#[test]
fn modulo_with_floats() {
    assert_eq!(run_print("7.5 % 2"), "1.5");
}

#[test]
fn divmod_unpack_in_assignment() {
    assert_eq!(run_python_one("q, r = divmod(17, 5)\nprint(q, r)\n"), "3 2");
}

#[test]
fn floordiv_augmented_assign() {
    assert_eq!(run_python_one("n = 17\nn //= 5\nprint(n)\n"), "3");
}

#[test]
fn modulo_augmented_assign() {
    assert_eq!(run_python_one("n = 17\nn %= 5\nprint(n)\n"), "2");
}

#[test]
fn divmod_zero_remainder_pattern() {
    assert_eq!(
        run_python_one("q, r = divmod(20, 4)\nprint(r == 0)\n"),
        "True"
    );
}

#[test]
fn floordiv_in_expression_with_add() {
    assert_eq!(run_print("(20 // 3) + (20 % 3)"), "8");
}

#[test]
fn modulo_one_always_zero() {
    assert_eq!(run_print("99 % 1"), "0");
}

#[test]
fn floordiv_self_by_self_is_one() {
    assert_eq!(run_print("15 // 15"), "1");
}

#[test]
fn divmod_self_by_self() {
    assert_eq!(run_print("divmod(15, 15)"), "(1, 0)");
}

#[test]
fn floordiv_smaller_than_divisor() {
    assert_eq!(run_print("2 // 10"), "0");
}

#[test]
fn modulo_smaller_than_divisor() {
    assert_eq!(run_print("2 % 10"), "2");
}

#[test]
fn divmod_smaller_than_divisor() {
    assert_eq!(run_print("divmod(2, 10)"), "(0, 2)");
}

#[test]
fn floordiv_in_list_comprehension() {
    assert_eq!(
        run_print("[n // 2 for n in [1, 2, 3, 4, 5]]"),
        "[0, 1, 1, 2, 2]"
    );
}

#[test]
fn modulo_in_list_comprehension() {
    assert_eq!(
        run_print("[n % 3 for n in range(7)]"),
        "[0, 1, 2, 0, 1, 2, 0]"
    );
}

#[test]
fn divmod_in_loop_accumulator() {
    assert_eq!(
        run_python_one(
            "total = 0\nfor n in [10, 11, 12]:\n q, r = divmod(n, 3)\n total += r\nprint(total)\n"
        ),
        "3"
    );
}

#[test]
fn floordiv_negative_float_dividend() {
    assert_eq!(run_print("-7.0 // 2"), "-4.0");
}

#[test]
fn modulo_preserves_divmod_identity() {
    assert_eq!(
        run_python_one("q, r = divmod(23, 4)\nprint(q * 4 + r)\n"),
        "23"
    );
}

#[test]
fn floordiv_chained_with_multiplication() {
    assert_eq!(run_print("(100 // 7) * 7"), "98");
}

#[test]
fn modulo_alternating_parity() {
    assert_eq!(run_print("[n % 2 for n in range(6)]"), "[0, 1, 0, 1, 0, 1]");
}

#[test]
fn divmod_with_negative_one_divisor() {
    assert_eq!(run_print("divmod(5, -1)"), "(-5, 0)");
}

#[test]
fn floordiv_with_negative_one_divisor() {
    assert_eq!(run_print("5 // -1"), "-5");
}

#[test]
fn modulo_with_negative_one_divisor() {
    assert_eq!(run_print("5 % -1"), "0");
}

#[test]
fn divmod_hour_minute_conversion() {
    assert_eq!(
        run_python_one("minutes = 125\nh, m = divmod(minutes, 60)\nprint(h, m)\n"),
        "2 5"
    );
}

#[test]
fn floordiv_bytes_per_kilobyte() {
    assert_eq!(run_print("2048 // 1024"), "2");
}

#[test]
fn modulo_circular_index_wrap() {
    assert_eq!(run_print("(7 % 5)"), "2");
}

#[test]
fn divmod_page_offset() {
    assert_eq!(run_print("divmod(47, 10)"), "(4, 7)");
}

#[test]
fn floordiv_half_of_even() {
    assert_eq!(run_print("100 // 2"), "50");
}

#[test]
fn floordiv_half_of_odd() {
    assert_eq!(run_print("101 // 2"), "50");
}

#[test]
fn modulo_last_digit_via_ten() {
    assert_eq!(run_print("12345 % 10"), "5");
}

#[test]
fn divmod_strips_last_digit() {
    assert_eq!(run_print("divmod(12345, 10)"), "(1234, 5)");
}

#[test]
fn floordiv_in_fstring() {
    assert_eq!(run_print("f'{17 // 5}'"), "3");
}

#[test]
fn modulo_in_fstring() {
    assert_eq!(run_print("f'{17 % 5}'"), "2");
}
