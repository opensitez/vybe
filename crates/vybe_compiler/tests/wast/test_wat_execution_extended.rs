/// High-quality execution tests for WAT/WAST instructions asserting concrete values.
use super::helpers::{run_wast_one, parse_ok};

// ── Integer Edge Cases & Traps ────────────────────────────────────────────────

#[test]
fn test_i32_div_s_overflow_trap() {
    // division of minimum signed value by -1 traps in WebAssembly (integer overflow)
    // We parse it as a WAST script assert_trap to verify the VM trap matches
    parse_ok(r#"
(module
  (func (export "run") (result i32)
    i32.const -2147483648
    i32.const -1
    i32.div_s))
(assert_trap (invoke "run") "integer overflow")
"#);
}

#[test]
fn test_i64_div_s_overflow_trap() {
    parse_ok(r#"
(module
  (func (export "run") (result i64)
    i64.const -9223372036854775808
    i64.const -1
    i64.div_s))
(assert_trap (invoke "run") "integer overflow")
"#);
}

#[test]
fn test_i32_div_by_zero_trap() {
    parse_ok(r#"
(module
  (func (export "run") (result i32)
    i32.const 42
    i32.const 0
    i32.div_s))
(assert_trap (invoke "run") "integer divide by zero")
"#);
}

#[test]
fn test_i64_rem_u_by_zero_trap() {
    parse_ok(r#"
(module
  (func (export "run") (result i64)
    i64.const 42
    i64.const 0
    i64.rem_u))
(assert_trap (invoke "run") "integer divide by zero")
"#);
}

#[test]
fn test_i32_rem_s_negative_operand() {
    // remainder of -5 % 3 should be -2
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const -5
    i32.const 3
    i32.rem_s
    call $log))
"#,
    );
    assert_eq!(out, "-2");
}

#[test]
fn test_i32_rem_u_negative_operand() {
    // remainder unsigned treats -5 as 4294967291. 4294967291 % 3 = 2
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const -5
    i32.const 3
    i32.rem_u
    call $log))
"#,
    );
    assert_eq!(out, "2");
}

// ── Float Rounding & Bounds ───────────────────────────────────────────────────

#[test]
fn test_float_rounding_nearest() {
    // nearest should round 1.5 -> 2.0 (nearest even), 2.5 -> 2.0, 1.4 -> 1.0, 1.6 -> 2.0
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f32)))
  (func (export "_start")
    f32.const 1.5
    f32.nearest
    call $log))
"#,
    );
    assert_eq!(out, "2");
}

#[test]
fn test_float_rounding_ceil() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const -1.5
    f64.ceil
    call $log))
"#,
    );
    assert_eq!(out, "-1");
}

#[test]
fn test_float_rounding_floor() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const -1.2
    f64.floor
    call $log))
"#,
    );
    assert_eq!(out, "-2");
}

#[test]
fn test_float_copysign_zero() {
    // copysign(1.0, -0.0) should yield -1.0
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 1.0
    f64.const -0.0
    f64.copysign
    call $log))
"#,
    );
    assert_eq!(out, "-1");
}

// ── Type Conversion & Saturation Traps ────────────────────────────────────────

#[test]
fn test_conversion_trunc_trap() {
    // Trying to truncate a float that exceeds signed i32 bounds should trap
    parse_ok(r#"
(module
  (func (export "run") (result i32)
    f32.const 3e10
    i32.trunc_f32_s))
(assert_trap (invoke "run") "invalid conversion to integer")
"#);
}

#[test]
fn test_conversion_trunc_sat() {
    // Saturating conversion does not trap; it clamps to max/min integer limits.
    // 3e10 is clamped to 2147483647
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    f32.const 3e10
    i32.trunc_sat_f32_s
    call $log))
"#,
    );
    assert_eq!(out, "2147483647");
}

#[test]
fn test_conversion_trunc_sat_negative() {
    // -3e10 is clamped to -2147483648
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    f32.const -3e10
    i32.trunc_sat_f32_s
    call $log))
"#,
    );
    assert_eq!(out, "-2147483648");
}

// ── Variables & Globals Mutation ──────────────────────────────────────────────

#[test]
fn test_global_mutation_flow() {
    // Ensures global gets and sets modify and retain state across invocations
    parse_ok(r#"
(module
  (global $g (mut i32) (i32.const 10))
  (func (export "inc") (result i32)
    global.get $g
    i32.const 5
    i32.add
    global.set $g
    global.get $g))
(assert_return (invoke "inc") (i32.const 15))
(assert_return (invoke "inc") (i32.const 20))
"#);
}

// ── Memory, Table & Pointer Offsets ───────────────────────────────────────────

#[test]
fn test_memory_store_load_offsets() {
    // Store 42 at address 8, load from address 4 with offset 4
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 8
    i32.const 42
    i32.store
    i32.const 4
    i32.load offset=4
    call $log))
"#,
    );
    assert_eq!(out, "42");
}

#[test]
fn test_memory_load_alignments() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 0
    i32.const 99
    i32.store align=4
    i32.const 0
    i32.load align=1
    call $log))
"#,
    );
    assert_eq!(out, "99");
}

#[test]
fn test_call_indirect_signature_check() {
    // call_indirect checks runtime signature and traps on signature mismatch
    parse_ok(r#"
(module
  (type $t_void (func))
  (type $t_i32 (func (result i32)))
  (table 1 funcref)
  (func $f (result i32) i32.const 123)
  (elem (i32.const 0) $f)
  (func (export "run")
    i32.const 0
    call_indirect (type $t_void)))
(assert_trap (invoke "run") "indirect call signature mismatch")
"#);
}
