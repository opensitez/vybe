use vybec::parser_ruby::parse;
use vybec::compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── Arithmetic ─────────────────────────────────────────────
#[test] fn add() { compile_ok("x = 1 + 2"); }
#[test] fn sub() { compile_ok("x = 5 - 3"); }
#[test] fn mul() { compile_ok("x = 4 * 3"); }
#[test] fn div() { compile_ok("x = 10 / 2"); }
#[test] fn modulo() { compile_ok("x = 10 % 3"); }
#[test] fn power() { compile_ok("x = 2 ** 10"); }

// ── Comparison ─────────────────────────────────────────────
#[test] fn eq() { compile_ok("x = 1 == 1"); }
#[test] fn ne() { compile_ok("x = 1 != 2"); }
#[test] fn lt() { compile_ok("x = 1 < 2"); }
#[test] fn gt() { compile_ok("x = 2 > 1"); }
#[test] fn le() { compile_ok("x = 1 <= 2"); }
#[test] fn ge() { compile_ok("x = 2 >= 1"); }
#[test] fn spaceship() { compile_ok("x = 1 <=> 2"); }

// ── Logical ────────────────────────────────────────────────
#[test] fn and_op() { compile_ok("x = true && false"); }
#[test] fn or_op() { compile_ok("x = true || false"); }
#[test] fn not_op() { compile_ok("x = !true"); }
#[test] fn and_word() { compile_ok("x = true and false"); }
#[test] fn or_word() { compile_ok("x = true or false"); }
#[test] fn not_word() { compile_ok("x = not true"); }

// ── Bitwise ────────────────────────────────────────────────
#[test] fn bit_and() { compile_ok("x = 5 & 3"); }
#[test] fn bit_or() { compile_ok("x = 5 | 3"); }
#[test] fn bit_xor() { compile_ok("x = 5 ^ 3"); }
#[test] fn bit_not() { compile_ok("x = ~5"); }
#[test] fn shift_left() { compile_ok("x = 1 << 3"); }
#[test] fn shift_right() { compile_ok("x = 8 >> 2"); }

// ── Assignment ─────────────────────────────────────────────
#[test] fn assign() { compile_ok("x = 5"); }
#[test] fn add_assign() { compile_ok("x = 1\nx += 2"); }
#[test] fn sub_assign() { compile_ok("x = 5\nx -= 3"); }
#[test] fn mul_assign() { compile_ok("x = 2\nx *= 3"); }
#[test] fn div_assign() { compile_ok("x = 10\nx /= 2"); }
#[test] fn mod_assign() { compile_ok("x = 10\nx %= 3"); }
#[test] fn and_assign() { compile_ok("x = true\nx &&= false"); }
#[test] fn or_assign() { compile_ok("x = nil\nx ||= 'default'"); }

// ── Ternary ────────────────────────────────────────────────
#[test] fn ternary() { compile_ok("x = true ? 'yes' : 'no'"); }

// ── String concatenation ──────────────────────────────────
#[test] fn str_concat() { compile_ok("x = 'hello' + ' ' + 'world'"); }
