use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// ── Operators (compile) ─────────────────────────────────────────────────────

#[test]
fn add() {
    compile_ok("x = 1 + 2\n");
}
#[test]
fn sub() {
    compile_ok("x = 5 - 3\n");
}
#[test]
fn mul() {
    compile_ok("x = 4 * 3\n");
}
#[test]
fn div() {
    compile_ok("x = 10 / 2\n");
}
#[test]
fn modulo() {
    compile_ok("x = 10 % 3\n");
}
#[test]
fn power() {
    compile_ok("x = 2 ** 10\n");
}
#[test]
fn eq() {
    compile_ok("x = 1 == 1\n");
}
#[test]
fn ne() {
    compile_ok("x = 1 != 2\n");
}
#[test]
fn lt() {
    compile_ok("x = 1 < 2\n");
}
#[test]
fn gt() {
    compile_ok("x = 2 > 1\n");
}
#[test]
fn le() {
    compile_ok("x = 1 <= 2\n");
}
#[test]
fn ge() {
    compile_ok("x = 2 >= 1\n");
}
#[test]
fn spaceship() {
    compile_ok("x = 1 <=> 2\n");
}
#[test]
fn and_op() {
    compile_ok("x = true && false\n");
}
#[test]
fn or_op() {
    compile_ok("x = true || false\n");
}
#[test]
fn not_op() {
    compile_ok("x = !true\n");
}
#[test]
fn and_word() {
    compile_ok("x = true and false\n");
}
#[test]
fn or_word() {
    compile_ok("x = true or false\n");
}
#[test]
fn not_word() {
    compile_ok("x = not true\n");
}
#[test]
fn bit_and() {
    compile_ok("x = 5 & 3\n");
}
#[test]
fn bit_or() {
    compile_ok("x = 5 | 3\n");
}
#[test]
fn bit_xor() {
    compile_ok("x = 5 ^ 3\n");
}
#[test]
fn bit_not() {
    compile_ok("x = ~5\n");
}
#[test]
fn shift_left() {
    compile_ok("x = 1 << 3\n");
}
#[test]
fn shift_right() {
    compile_ok("x = 8 >> 2\n");
}
#[test]
fn assign() {
    compile_ok("x = 5\n");
}
#[test]
fn ternary() {
    compile_ok("x = true ? 'yes' : 'no'\n");
}
#[test]
fn str_concat() {
    compile_ok("x = 'hello' + ' ' + 'world'\n");
}

// ── Augmented assignment ────────────────────────────────────────────────────

#[test]
fn add_assign() {
    compile_ok("x = 1\nx += 2\n");
}
#[test]
fn sub_assign() {
    compile_ok("x = 5\nx -= 3\n");
}
#[test]
fn mul_assign() {
    compile_ok("x = 2\nx *= 3\n");
}
#[test]
fn div_assign() {
    compile_ok("x = 10\nx /= 2\n");
}
#[test]
fn mod_assign() {
    compile_ok("x = 10\nx %= 3\n");
}
#[test]
fn and_assign() {
    compile_ok("x = true\nx &&= false\n");
}
#[test]
fn or_assign() {
    compile_ok("x = nil\nx ||= 'default'\n");
}

// ── Runtime ─────────────────────────────────────────────────────────────────

#[test]
fn add_runtime() {
    assert_eq!(run_ruby_one("puts 1 + 2\n"), "3");
}

#[test]
fn sub_runtime() {
    assert_eq!(run_ruby_one("puts 5 - 3\n"), "2");
}

#[test]
fn mul_runtime() {
    assert_eq!(run_ruby_one("puts 4 * 3\n"), "12");
}

#[test]
fn div_runtime() {
    assert_eq!(run_ruby_one("puts 10 / 2\n"), "5");
}

#[test]
fn modulo_runtime() {
    assert_eq!(run_ruby_one("puts 10 % 3\n"), "1");
}

#[test]
fn power_runtime() {
    assert_eq!(run_ruby_one("puts 2 ** 10\n"), "1024");
}

#[test]
fn comparison_runtime() {
    assert_eq!(run_ruby_one("puts 1 < 2\n"), "true");
}

#[test]
fn bool_and_runtime() {
    assert_eq!(run_ruby_one("puts true && false\n"), "false");
}

#[test]
fn augmented_assign_runtime() {
    let out = run_ruby("x = 10\nx += 5\nputs x\n");
    assert_eq!(out, vec!["15"]);
}

#[test]
fn ternary_runtime() {
    assert_eq!(run_ruby_one("puts(true ? 'yes' : 'no')\n"), "yes");
}
