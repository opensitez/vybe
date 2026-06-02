/// Tests for WAT folded (S-expression) instruction syntax
use super::helpers::{compile_ok, parse_ok};

// ── Basic folded instructions ─────────────────────────────────────────────────

#[test]
fn folded_const() {
    parse_ok("(module (func (result i32) (i32.const 42)))");
}

#[test]
fn folded_add_consts() {
    parse_ok("(module (func (result i32) (i32.add (i32.const 1) (i32.const 2))))");
}

#[test]
fn folded_local_get() {
    parse_ok("(module (func (param $x i32) (result i32) (local.get $x)))");
}

#[test]
fn folded_local_set() {
    parse_ok("(module (func (param i32) (local i32) (local.set 1 (local.get 0))))");
}

#[test]
fn folded_local_tee() {
    parse_ok("(module (func (param i32) (result i32) (local i32) (local.tee 1 (local.get 0))))");
}

#[test]
fn folded_global_get() {
    parse_ok("(module (global $g i32 (i32.const 0)) (func (result i32) (global.get $g)))");
}

#[test]
fn folded_global_set() {
    parse_ok("(module (global $g (mut i32) (i32.const 0)) (func (global.set $g (i32.const 1))))");
}

// ── Folded arithmetic ─────────────────────────────────────────────────────────

#[test]
fn folded_sub() {
    parse_ok("(module (func (param i32 i32) (result i32) (i32.sub (local.get 0) (local.get 1))))");
}

#[test]
fn folded_mul() {
    parse_ok("(module (func (param i32 i32) (result i32) (i32.mul (local.get 0) (local.get 1))))");
}

#[test]
fn folded_div_s() {
    parse_ok(
        "(module (func (param i32 i32) (result i32) (i32.div_s (local.get 0) (local.get 1))))",
    );
}

#[test]
fn folded_rem_u() {
    parse_ok(
        "(module (func (param i32 i32) (result i32) (i32.rem_u (local.get 0) (local.get 1))))",
    );
}

#[test]
fn folded_f64_mul() {
    parse_ok("(module (func (param f64 f64) (result f64) (f64.mul (local.get 0) (local.get 1))))");
}

#[test]
fn folded_f64_sqrt() {
    parse_ok("(module (func (param f64) (result f64) (f64.sqrt (local.get 0))))");
}

// ── Folded comparisons ────────────────────────────────────────────────────────

#[test]
fn folded_i32_eq() {
    parse_ok("(module (func (param i32 i32) (result i32) (i32.eq (local.get 0) (local.get 1))))");
}

#[test]
fn folded_i32_lt_s() {
    parse_ok("(module (func (param i32 i32) (result i32) (i32.lt_s (local.get 0) (local.get 1))))");
}

#[test]
fn folded_i32_eqz() {
    parse_ok("(module (func (param i32) (result i32) (i32.eqz (local.get 0))))");
}

// ── Folded control flow ───────────────────────────────────────────────────────

#[test]
fn folded_call() {
    parse_ok(
        "(module (func $f (param i32) (result i32) (local.get 0)) (func (result i32) (call $f (i32.const 5))))",
    );
}

#[test]
fn folded_return() {
    parse_ok("(module (func (result i32) (return (i32.const 42))))");
}

#[test]
fn folded_drop() {
    parse_ok("(module (func (drop (i32.const 1))))");
}

#[test]
fn folded_select() {
    parse_ok(
        "(module (func (param i32 i32 i32) (result i32) (select (local.get 0) (local.get 1) (local.get 2))))",
    );
}

#[test]
fn folded_br() {
    parse_ok("(module (func (block $b (br $b))))");
}

#[test]
fn folded_br_if() {
    parse_ok("(module (func (param i32) (block $b (br_if $b (local.get 0)))))");
}

// ── Deeply nested folded ──────────────────────────────────────────────────────

#[test]
fn folded_three_levels() {
    parse_ok(
        "(module (func (param i32 i32 i32) (result i32) (i32.add (i32.mul (local.get 0) (local.get 1)) (local.get 2))))",
    );
}

#[test]
fn folded_four_levels() {
    parse_ok(
        "(module (func (param i32) (result i32) (i32.add (i32.mul (i32.add (local.get 0) (i32.const 1)) (i32.const 2)) (i32.const 3))))",
    );
}

#[test]
fn folded_mixed_with_plain() {
    parse_ok(
        r#"
(module
  (func (param $a i32) (param $b i32) (result i32)
    (i32.add (local.get $a) (local.get $b))
    local.get $a
    i32.add))
"#,
    );
}

// ── Folded memory ─────────────────────────────────────────────────────────────

#[test]
fn folded_i32_load() {
    parse_ok("(module (memory 1) (func (param i32) (result i32) (i32.load (local.get 0))))");
}

#[test]
fn folded_i32_store() {
    parse_ok("(module (memory 1) (func (param i32 i32) (i32.store (local.get 0) (local.get 1))))");
}

#[test]
fn folded_memory_grow() {
    parse_ok("(module (memory 1) (func (param i32) (result i32) (memory.grow (local.get 0))))");
}

// ── Folded type conversions ───────────────────────────────────────────────────

#[test]
fn folded_i64_extend_i32_s() {
    parse_ok("(module (func (param i32) (result i64) (i64.extend_i32_s (local.get 0))))");
}

#[test]
fn folded_f64_convert_i32_s() {
    parse_ok("(module (func (param i32) (result f64) (f64.convert_i32_s (local.get 0))))");
}

#[test]
fn folded_i32_wrap_i64() {
    parse_ok("(module (func (param i64) (result i32) (i32.wrap_i64 (local.get 0))))");
}

// ── Compile checks ────────────────────────────────────────────────────────────

#[test]
fn compile_folded_factorial() {
    compile_ok(
        r#"
(module
  (func $fact (export "fact") (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
      (then (i32.const 1))
      (else (i32.mul (local.get $n)
                     (call $fact (i32.sub (local.get $n) (i32.const 1))))))))
"#,
    );
}

#[test]
fn compile_folded_min() {
    compile_ok(
        r#"
(module
  (func $min (export "min") (param $a i32) (param $b i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $a) (local.get $b))
      (then (local.get $a))
      (else (local.get $b)))))
"#,
    );
}
