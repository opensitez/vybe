use super::helpers::*;

// ══════════════════════════════════════════════════════════════════════════════
// Arithmetic operators runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn floor_division() {
    assert_eq!(run_python_one("print(10 // 3)\n"), "3");
}

#[test]
fn modulo() {
    assert_eq!(run_python_one("print(10 % 3)\n"), "1");
}

#[test]
fn power_operator() {
    assert_eq!(run_python_one("print(2 ** 10)\n"), "1024");
}

#[test]
fn negative_floor_division() {
    assert_eq!(run_python_one("print(-7 // 2)\n"), "-4");
}

// ══════════════════════════════════════════════════════════════════════════════
// Bitwise operators runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn bitwise_and() {
    assert_eq!(run_python_one("print(5 & 3)\n"), "1");
}

#[test]
fn bitwise_or() {
    assert_eq!(run_python_one("print(5 | 3)\n"), "7");
}

#[test]
fn bitwise_xor() {
    assert_eq!(run_python_one("print(5 ^ 3)\n"), "6");
}

#[test]
fn bitwise_not() {
    assert_eq!(run_python_one("print(~5)\n"), "-6");
}

#[test]
fn shift_left() {
    assert_eq!(run_python_one("print(1 << 4)\n"), "16");
}

#[test]
fn shift_right() {
    assert_eq!(run_python_one("print(16 >> 2)\n"), "4");
}

// ══════════════════════════════════════════════════════════════════════════════
// Unary operators runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn unary_negative() {
    assert_eq!(run_python_one("print(-5)\n"), "-5");
}

#[test]
fn unary_positive() {
    assert_eq!(run_python_one("print(+5)\n"), "5");
}

// ══════════════════════════════════════════════════════════════════════════════
// Boolean operators runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn bool_and() {
    assert_eq!(run_python_one("print(True and False)\n"), "false");
}

#[test]
fn bool_or() {
    assert_eq!(run_python_one("print(True or False)\n"), "true");
}

#[test]
fn bool_not() {
    assert_eq!(run_python_one("print(not True)\n"), "false");
}

#[test]
fn bool_not_false() {
    assert_eq!(run_python_one("print(not False)\n"), "true");
}

// ══════════════════════════════════════════════════════════════════════════════
// Ternary expression runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn ternary_true_branch() {
    assert_eq!(run_python_one("x = 'yes' if True else 'no'\nprint(x)\n"), "yes");
}

#[test]
fn ternary_false_branch() {
    assert_eq!(run_python_one("x = 'yes' if False else 'no'\nprint(x)\n"), "no");
}

#[test]
fn ternary_with_expression() {
    assert_eq!(run_python_one("n = 7\nresult = 'even' if n % 2 == 0 else 'odd'\nprint(result)\n"), "odd");
}

// ══════════════════════════════════════════════════════════════════════════════
// Comparison operators runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn comparison_less_than() {
    assert_eq!(run_python_one("print(1 < 2)\n"), "true");
}

#[test]
fn comparison_greater_equal() {
    assert_eq!(run_python_one("print(5 >= 5)\n"), "true");
}

#[test]
fn comparison_not_equal() {
    assert_eq!(run_python_one("print(1 != 2)\n"), "true");
}

#[test]
fn in_operator() {
    assert_eq!(run_python_one("print(2 in [1, 2, 3])\n"), "true");
}

#[test]
fn not_in_operator() {
    assert_eq!(run_python_one("print(5 not in [1, 2, 3])\n"), "true");
}

#[test]
fn is_none() {
    compile_ok("x = None\nprint(x is None)\n");
}

#[test]
fn is_not_none() {
    compile_ok("x = 42\nprint(x is not None)\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Chained comparisons
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn chained_comparison_parse() {
    compile_ok("x = 1 < 2 < 3\n");
}

#[test]
fn chained_range_check() {
    compile_ok("x = 0 <= value <= 100\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// All augmented assignments
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn aug_assign_div() {
    assert_eq!(run_python_one("x = 10\nx /= 2\nprint(x)\n"), "5");
}

#[test]
fn aug_assign_mod() {
    assert_eq!(run_python_one("x = 10\nx %= 3\nprint(x)\n"), "1");
}

#[test]
fn aug_assign_pow() {
    assert_eq!(run_python_one("x = 2\nx **= 10\nprint(x)\n"), "1024");
}

#[test]
fn aug_assign_bitand() {
    compile_ok("x = 0xFF\nx &= 0x0F\n");
}

#[test]
fn aug_assign_bitor() {
    compile_ok("x = 0x0F\nx |= 0xF0\n");
}

#[test]
fn aug_assign_bitxor() {
    compile_ok("x = 0xFF\nx ^= 0x0F\n");
}

#[test]
fn aug_assign_shl() {
    compile_ok("x = 1\nx <<= 4\n");
}

#[test]
fn aug_assign_shr() {
    compile_ok("x = 16\nx >>= 2\n");
}

#[test]
fn aug_assign_floordiv() {
    assert_eq!(run_python_one("x = 10\nx //= 3\nprint(x)\n"), "3");
}
