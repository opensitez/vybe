/// Tests for WAST script commands — invoke, assert_return, assert_trap,
/// assert_invalid, assert_malformed, assert_unlinkable, register, get
use super::helpers::{compile_ok, parse_ok};

// ── invoke ────────────────────────────────────────────────────────────────────

#[test]
fn invoke_no_args() {
    parse_ok(
        r#"
(module (func (export "noop")))
(invoke "noop")
"#,
    );
}

#[test]
fn invoke_i32_arg() {
    parse_ok(
        r#"
(module (func (export "f") (param i32)))
(invoke "f" (i32.const 42))
"#,
    );
}

#[test]
fn invoke_multiple_args() {
    parse_ok(
        r#"
(module (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
(invoke "add" (i32.const 3) (i32.const 4))
"#,
    );
}

#[test]
fn invoke_f32_arg() {
    parse_ok(
        r#"
(module (func (export "f") (param f32)))
(invoke "f" (f32.const 1.5))
"#,
    );
}

#[test]
fn invoke_f64_arg() {
    parse_ok(
        r#"
(module (func (export "f") (param f64)))
(invoke "f" (f64.const 2.718))
"#,
    );
}

#[test]
fn invoke_i64_arg() {
    parse_ok(
        r#"
(module (func (export "f") (param i64)))
(invoke "f" (i64.const 9999999999))
"#,
    );
}

#[test]
fn invoke_on_named_module() {
    parse_ok(
        r#"
(module $m (func (export "f")))
(invoke $m "f")
"#,
    );
}

// ── assert_return ─────────────────────────────────────────────────────────────

#[test]
fn assert_return_i32() {
    parse_ok(
        r#"
(module (func (export "forty_two") (result i32) i32.const 42))
(assert_return (invoke "forty_two") (i32.const 42))
"#,
    );
}

#[test]
fn assert_return_i64() {
    parse_ok(
        r#"
(module (func (export "f") (result i64) i64.const 100))
(assert_return (invoke "f") (i64.const 100))
"#,
    );
}

#[test]
fn assert_return_f32() {
    parse_ok(
        r#"
(module (func (export "f") (result f32) f32.const 1.0))
(assert_return (invoke "f") (f32.const 1.0))
"#,
    );
}

#[test]
fn assert_return_f64() {
    parse_ok(
        r#"
(module (func (export "f") (result f64) f64.const 3.14))
(assert_return (invoke "f") (f64.const 3.14))
"#,
    );
}

#[test]
fn assert_return_no_result() {
    parse_ok(
        r#"
(module (func (export "noop")))
(assert_return (invoke "noop"))
"#,
    );
}

#[test]
fn assert_return_with_args() {
    parse_ok(
        r#"
(module (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
(assert_return (invoke "add" (i32.const 3) (i32.const 4)) (i32.const 7))
"#,
    );
}

#[test]
fn assert_return_multiple_results() {
    parse_ok(
        r#"
(module (func (export "f") (result i32 i32) i32.const 1 i32.const 2))
(assert_return (invoke "f") (i32.const 1) (i32.const 2))
"#,
    );
}

#[test]
fn assert_return_nan_canonical() {
    parse_ok(
        r#"
(module (func (export "f") (result f32) f32.const nan))
(assert_return (invoke "f") (f32.const nan:canonical))
"#,
    );
}

#[test]
fn assert_return_nan_arithmetic() {
    parse_ok(
        r#"
(module (func (export "f") (result f64) f64.const nan))
(assert_return (invoke "f") (f64.const nan:arithmetic))
"#,
    );
}

#[test]
fn assert_return_ref_null() {
    parse_ok(
        r#"
(module (func (export "f") (result funcref) ref.null func))
(assert_return (invoke "f") (ref.null func))
"#,
    );
}

// ── assert_trap ───────────────────────────────────────────────────────────────

#[test]
fn assert_trap_unreachable() {
    parse_ok(
        r#"
(module (func (export "f") unreachable))
(assert_trap (invoke "f") "unreachable")
"#,
    );
}

#[test]
fn assert_trap_div_zero() {
    parse_ok(
        r#"
(module (func (export "div") (param i32 i32) (result i32) local.get 0 local.get 1 i32.div_s))
(assert_trap (invoke "div" (i32.const 1) (i32.const 0)) "integer divide by zero")
"#,
    );
}

#[test]
fn assert_trap_oob_memory() {
    parse_ok(
        r#"
(module (memory 1) (func (export "load") (param i32) (result i32) local.get 0 i32.load))
(assert_trap (invoke "load" (i32.const 65536)) "out of bounds memory access")
"#,
    );
}

#[test]
fn assert_trap_stack_overflow() {
    parse_ok(
        r#"
(module (func $rec (export "rec") call $rec))
(assert_trap (invoke "rec") "call stack exhausted")
"#,
    );
}

// ── assert_exhaustion ─────────────────────────────────────────────────────────

#[test]
fn assert_exhaustion() {
    parse_ok(
        r#"
(module (func $inf (export "inf") call $inf))
(assert_exhaustion (invoke "inf") "call stack exhausted")
"#,
    );
}

// ── assert_invalid ────────────────────────────────────────────────────────────

#[test]
fn assert_invalid_type_mismatch() {
    // The outer WAST script parses fine; the inner module is intentionally invalid WAT
    // (type mismatch) which assert_invalid is designed to test.
    parse_ok(r#"(assert_invalid (module (func (result i32) f32.const 1.0)) "type mismatch")"#);
}

#[test]
fn assert_invalid_unknown_local() {
    parse_ok(r#"(assert_invalid (module (func (result i32) i32.const 1)) "unknown local")"#);
}

#[test]
fn assert_invalid_empty_stack() {
    parse_ok(r#"(assert_invalid (module (func (result i32) nop)) "type mismatch")"#);
}

// ── assert_malformed ──────────────────────────────────────────────────────────

#[test]
fn assert_malformed_binary() {
    parse_ok(r#"(assert_malformed (module binary "\00asm") "magic header not detected")"#);
}

#[test]
fn assert_malformed_quote() {
    parse_ok(
        r#"(assert_malformed (module quote "(module (func (result i32)))") "unexpected token")"#,
    );
}

// ── assert_unlinkable ─────────────────────────────────────────────────────────

#[test]
fn assert_unlinkable_missing_import() {
    parse_ok(r#"(assert_unlinkable (module (import "env" "missing" (func))) "unknown import")"#);
}

// ── register ─────────────────────────────────────────────────────────────────

#[test]
fn register_module() {
    parse_ok(
        r#"
(module $m (func (export "f") (result i32) i32.const 1))
(register "mymod" $m)
"#,
    );
}

#[test]
fn register_anonymous_module() {
    parse_ok(
        r#"
(module (func (export "f") (result i32) i32.const 1))
(register "mymod")
"#,
    );
}

// ── get ───────────────────────────────────────────────────────────────────────

#[test]
fn get_global() {
    parse_ok(
        r#"
(module (global (export "g") i32 (i32.const 42)))
(assert_return (get "g") (i32.const 42))
"#,
    );
}

#[test]
fn get_global_from_named_module() {
    parse_ok(
        r#"
(module $m (global (export "g") i32 (i32.const 7)))
(assert_return (get $m "g") (i32.const 7))
"#,
    );
}

// ── Multiple modules in one script ───────────────────────────────────────────

#[test]
fn multiple_modules() {
    parse_ok(
        r#"
(module $a (func (export "f") (result i32) i32.const 1))
(module $b (func (export "g") (result i32) i32.const 2))
(assert_return (invoke $a "f") (i32.const 1))
(assert_return (invoke $b "g") (i32.const 2))
"#,
    );
}

#[test]
fn module_then_assertions() {
    parse_ok(
        r#"
(module
  (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add)
  (func (export "mul") (param i32 i32) (result i32) local.get 0 local.get 1 i32.mul)
)
(assert_return (invoke "add" (i32.const 2) (i32.const 3)) (i32.const 5))
(assert_return (invoke "mul" (i32.const 4) (i32.const 5)) (i32.const 20))
(assert_return (invoke "add" (i32.const 0) (i32.const 0)) (i32.const 0))
"#,
    );
}

// ── Compile checks ────────────────────────────────────────────────────────────

#[test]
fn compile_assert_return() {
    compile_ok(
        r#"
(module (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
(assert_return (invoke "add" (i32.const 1) (i32.const 2)) (i32.const 3))
"#,
    );
}

#[test]
fn compile_assert_trap() {
    compile_ok(
        r#"
(module (func (export "boom") unreachable))
(assert_trap (invoke "boom") "unreachable")
"#,
    );
}

#[test]
fn compile_register_and_invoke() {
    compile_ok(
        r#"
(module $lib (func (export "double") (param i32) (result i32) local.get 0 i32.const 2 i32.mul))
(register "lib")
(invoke "double" (i32.const 5))
"#,
    );
}
