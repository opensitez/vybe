use super::helpers::{compile_ok, run_prints};

// ── Arithmetic ──────────────────────────────────────────────

#[test]
fn integer_division() {
    compile_ok("var x = 7 ~/ 2;");
}
#[test]
fn integer_division_result() {
    let out = run_prints("void main() { print(7 ~/ 2); }");
    assert_eq!(out, ["3"]);
}

#[test]
fn modulo() {
    compile_ok("var x = 10 % 3;");
}
#[test]
fn modulo_result() {
    let out = run_prints("void main() { print(10 % 3); }");
    assert_eq!(out, ["1"]);
}

#[test]
fn unary_minus() {
    compile_ok("var x = -5;");
}
#[test]
fn unary_minus_result() {
    let out = run_prints("void main() { var x = 5; print(-x); }");
    assert_eq!(out, ["-5"]);
}

// ── Compound assignment ──────────────────────────────────────

#[test]
fn plus_assign() {
    compile_ok("void main() { var x = 1; x += 2; print(x); }");
}
#[test]
fn minus_assign() {
    compile_ok("void main() { var x = 5; x -= 3; print(x); }");
}
#[test]
fn times_assign() {
    compile_ok("void main() { var x = 4; x *= 3; print(x); }");
}
#[test]
fn divide_assign() {
    compile_ok("void main() { var x = 10.0; x /= 2; print(x); }");
}
#[test]
fn int_div_assign() {
    compile_ok("void main() { var x = 9; x ~/= 2; print(x); }");
}
#[test]
fn mod_assign() {
    compile_ok("void main() { var x = 10; x %= 3; print(x); }");
}

#[test]
fn plus_assign_result() {
    let out = run_prints("void main() { var x = 10; x += 5; print(x); }");
    assert_eq!(out, ["15"]);
}

#[test]
fn minus_assign_result() {
    let out = run_prints("void main() { var x = 10; x -= 3; print(x); }");
    assert_eq!(out, ["7"]);
}

#[test]
fn times_assign_result() {
    let out = run_prints("void main() { var x = 6; x *= 7; print(x); }");
    assert_eq!(out, ["42"]);
}

// ── Increment / decrement ────────────────────────────────────

#[test]
fn pre_increment() {
    compile_ok("void main() { var x = 0; ++x; print(x); }");
}
#[test]
fn post_increment() {
    compile_ok("void main() { var x = 0; x++; print(x); }");
}
#[test]
fn pre_decrement() {
    compile_ok("void main() { var x = 5; --x; print(x); }");
}
#[test]
fn post_decrement() {
    compile_ok("void main() { var x = 5; x--; print(x); }");
}

#[test]
fn pre_increment_result() {
    let out = run_prints("void main() { var x = 0; ++x; print(x); }");
    assert_eq!(out, ["1"]);
}

#[test]
fn post_increment_in_loop() {
    let out =
        run_prints("void main() { var x = 0; for (var i = 0; i < 3; i++) { x++; } print(x); }");
    assert_eq!(out, ["3"]);
}

#[test]
fn decrement_result() {
    let out = run_prints("void main() { var x = 10; x--; x--; print(x); }");
    assert_eq!(out, ["8"]);
}

// ── Bitwise ─────────────────────────────────────────────────

#[test]
fn bitwise_and() {
    compile_ok("var x = 0xFF & 0x0F;");
}
#[test]
fn bitwise_or() {
    compile_ok("var x = 0xF0 | 0x0F;");
}
#[test]
fn bitwise_xor() {
    compile_ok("var x = 0xFF ^ 0x0F;");
}
#[test]
fn bitwise_not() {
    compile_ok("var x = ~0;");
}
#[test]
fn left_shift() {
    compile_ok("var x = 1 << 4;");
}
#[test]
fn right_shift() {
    compile_ok("var x = 256 >> 4;");
}

#[test]
fn bitwise_and_result() {
    let out = run_prints("void main() { print(0xFF & 0x0F); }");
    assert_eq!(out, ["15"]);
}

#[test]
fn bitwise_or_result() {
    let out = run_prints("void main() { print(0xF0 | 0x0F); }");
    assert_eq!(out, ["255"]);
}

#[test]
fn left_shift_result() {
    let out = run_prints("void main() { print(1 << 3); }");
    assert_eq!(out, ["8"]);
}

#[test]
fn right_shift_result() {
    let out = run_prints("void main() { print(64 >> 2); }");
    assert_eq!(out, ["16"]);
}

#[test]
fn bitwise_assign() {
    compile_ok("void main() { var x = 0xFF; x &= 0x0F; print(x); }");
}
#[test]
fn bitor_assign() {
    compile_ok("void main() { var x = 0xF0; x |= 0x0F; print(x); }");
}
#[test]
fn xor_assign() {
    compile_ok("void main() { var x = 0xFF; x ^= 0x0F; print(x); }");
}
#[test]
fn shift_assign() {
    compile_ok("void main() { var x = 1; x <<= 3; print(x); }");
}

// ── Boolean operators ────────────────────────────────────────

#[test]
fn logical_and() {
    compile_ok("var x = true && false;");
}
#[test]
fn logical_or() {
    compile_ok("var x = false || true;");
}
#[test]
fn logical_not() {
    compile_ok("var x = !true;");
}
#[test]
fn logical_not_result() {
    let out = run_prints("void main() { print(!false); }");
    assert_eq!(out, ["true"]);
}

#[test]
fn short_circuit_and() {
    let out = run_prints("void main() { var x = 0; var y = false && (x = 1) == 1; print(x); }");
    assert_eq!(out, ["0"]);
}

// ── Ternary ─────────────────────────────────────────────────

#[test]
fn ternary_basic() {
    compile_ok("var x = true ? 'yes' : 'no';");
}
#[test]
fn ternary_result() {
    let out = run_prints("void main() { var x = 5 > 3 ? 'big' : 'small'; print(x); }");
    assert_eq!(out, ["big"]);
}

#[test]
fn ternary_nested() {
    compile_ok("var x = 5; var label = x > 10 ? 'large' : x > 5 ? 'medium' : 'small';");
}

#[test]
fn ternary_in_string_interp() {
    compile_ok("var n = 3; var s = 'There are ${n == 1 ? \"item\" : \"items\"}';");
}

// ── Null-coalescing assignment ───────────────────────────────

#[test]
fn null_coalesce_assign() {
    compile_ok("var x; x ??= 42;");
}
#[test]
fn null_coalesce_assign_result() {
    let out = run_prints("void main() { var x; x ??= 42; print(x); }");
    assert_eq!(out, ["42"]);
}

#[test]
fn null_coalesce_no_overwrite() {
    let out = run_prints("void main() { var x = 10; x ??= 99; print(x); }");
    assert_eq!(out, ["10"]);
}

// ── Comparison ──────────────────────────────────────────────

#[test]
fn eq_operator() {
    compile_ok("var x = 1 == 1;");
}
#[test]
fn neq_operator() {
    compile_ok("var x = 1 != 2;");
}
#[test]
fn lt_operator() {
    compile_ok("var x = 1 < 2;");
}
#[test]
fn gt_operator() {
    compile_ok("var x = 2 > 1;");
}
#[test]
fn lte_operator() {
    compile_ok("var x = 1 <= 1;");
}
#[test]
fn gte_operator() {
    compile_ok("var x = 2 >= 2;");
}

#[test]
fn comparison_chain() {
    let out = run_prints("void main() { var a = 5; print(a >= 1 && a <= 10); }");
    assert_eq!(out, ["true"]);
}

// ── String concatenation ─────────────────────────────────────

#[test]
fn string_concat_plus() {
    compile_ok("var s = 'Hello' + ' ' + 'World';");
}
#[test]
fn string_concat_result() {
    let out = run_prints("void main() { var s = 'foo' + 'bar'; print(s); }");
    assert_eq!(out, ["foobar"]);
}

#[test]
fn string_concat_assign() {
    compile_ok("void main() { var s = 'Hello'; s += ' World'; print(s); }");
}
