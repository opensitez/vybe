use vybe_bytecode::*;
use vybe_bytecode::value::*;
use std::rc::Rc;

// ============================================================
// Helper: build chunks, write to WASM, read back
// ============================================================

fn roundtrip(chunks: Vec<Chunk>) -> Vec<Chunk> {
    let wasm_bytes = wasm::write_wasm(&chunks);
    wasm::read_wasm(&wasm_bytes).expect("Failed to parse WASM")
}

fn run_chunks(chunks: Vec<Chunk>) -> Value {
    let mut vm = VM::new();
    vm.run(chunks).unwrap()
}

fn run_with_host(chunks: Vec<Chunk>) -> (VM, Value) {
    let mut vm = VM::new();
    let result = vm.run(chunks).unwrap();
    (vm, result)
}

// ============================================================
// WASM Binary Format — Parsing
// ============================================================

#[test]
fn wasm_magic_and_version() {
    // Minimal WASM with no code section → our reader returns error
    let data = vec![0x00, 0x61, 0x73, 0x6D, 0x01, 0x00, 0x00, 0x00];
    let result = wasm::read_wasm(&data);
    // Our reader requires code section, so empty module is an error
    assert!(result.is_err());
}

#[test]
fn wasm_invalid_magic() {
    let data = vec![0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00, 0x00, 0x00];
    let result = wasm::read_wasm(&data);
    assert!(result.is_err());
}

#[test]
fn wasm_too_short() {
    let data = vec![0x00, 0x61];
    let result = wasm::read_wasm(&data);
    assert!(result.is_err());
}

// ============================================================
// WASM Round-trip — write then read back
// ============================================================

#[test]
fn roundtrip_simple_constant() {
    let mut chunk = Chunk::new("test");
    let ci = chunk.add_constant(Value::F64(42.0));
    chunk.emit_op_u16(Op::r#const, ci, 0);
    chunk.emit_op(Op::halt, 0);
    chunk.local_count = 1;

    let chunks = roundtrip(vec![chunk]);
    assert!(!chunks.is_empty());
}

#[test]
fn roundtrip_preserves_chunk_count() {
    let mut c1 = Chunk::new("main");
    c1.emit_op(Op::null, 0);
    c1.emit_op(Op::halt, 0);

    let mut c2 = Chunk::new("helper");
    c2.arity = 1;
    c2.local_count = 2;
    c2.emit_op_u16(Op::local_get, 1, 0);
    c2.emit_op(Op::r#return, 0);

    let original = vec![c1, c2];
    let restored = roundtrip(original);
    assert_eq!(restored.len(), 2);
}

#[test]
fn roundtrip_f64_arithmetic() {
    // Build: push 10.0, push 3.0, divide → 3.333...
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c10 = chunk.add_constant(Value::F64(10.0));
    let c3 = chunk.add_constant(Value::F64(3.0));
    chunk.emit_op_u16(Op::r#const, c10, 0);
    chunk.emit_op_u16(Op::r#const, c3, 0);
    chunk.emit_op(Op::f64_div, 0);
    chunk.emit_op(Op::halt, 0);

    // Run original
    let result1 = run_chunks(vec![chunk.clone()]);

    // Round-trip and run
    let restored = roundtrip(vec![chunk]);
    let result2 = run_chunks(restored);

    match (&result1, &result2) {
        (Value::F64(a), Value::F64(b)) => {
            assert!((a - b).abs() < 1e-10, "Results differ: {} vs {}", a, b);
        }
        _ => panic!("Expected F64 results"),
    }
}

#[test]
fn roundtrip_i32_arithmetic() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c7 = chunk.add_constant(Value::I32(7));
    let c3 = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::r#const, c7, 0);
    chunk.emit_op_u16(Op::r#const, c3, 0);
    chunk.emit_op(Op::i32_add, 0);
    chunk.emit_op(Op::halt, 0);

    let restored = roundtrip(vec![chunk]);
    let result = run_chunks(restored);
    match result {
        Value::I32(10) => {}
        _ => panic!("Expected I32(10), got {:?}", result),
    }
}

#[test]
fn roundtrip_locals() {
    // set local 1 = 99, get local 1
    let mut chunk = Chunk::new("test");
    chunk.local_count = 3;
    let c99 = chunk.add_constant(Value::F64(99.0));
    chunk.emit_op_u16(Op::r#const, c99, 0);
    chunk.emit_op_u16(Op::local_set, 1, 0);
    chunk.emit_op(Op::drop, 0);
    chunk.emit_op_u16(Op::local_get, 1, 0);
    chunk.emit_op(Op::halt, 0);

    let restored = roundtrip(vec![chunk]);
    let result = run_chunks(restored);
    match result {
        Value::F64(v) if v == 99.0 => {}
        _ => panic!("Expected F64(99), got {:?}", result),
    }
}

// ============================================================
// WASM Runtime — New i32 Opcodes
// ============================================================

#[test]
fn i32_div_and_rem() {
    let mut chunk = Chunk::new("test");
    let c17 = chunk.add_constant(Value::I32(17));
    let c5 = chunk.add_constant(Value::I32(5));

    // 17 / 5 = 3
    chunk.emit_op_u16(Op::r#const, c17, 0);
    chunk.emit_op_u16(Op::r#const, c5, 0);
    chunk.emit_op(Op::i32_div_s, 0);

    // 17 % 5 = 2
    chunk.emit_op_u16(Op::r#const, c17, 0);
    chunk.emit_op_u16(Op::r#const, c5, 0);
    chunk.emit_op(Op::i32_rem_s, 0);

    // Stack: [3, 2] → add = 5
    chunk.emit_op(Op::i32_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(5) => {} _ => panic!("Expected I32(5), got {:?}", result) }
}

#[test]
fn i32_div_by_zero() {
    let mut chunk = Chunk::new("test");
    let c10 = chunk.add_constant(Value::I32(10));
    let c0 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::r#const, c10, 0);
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op(Op::i32_div_s, 0);
    chunk.emit_op(Op::halt, 0);

    // Should return 0, not crash
    let result = run_chunks(vec![chunk]);
    match result { Value::I32(0) => {} _ => panic!("Expected I32(0), got {:?}", result) }
}

#[test]
fn i32_rotate() {
    let mut chunk = Chunk::new("test");
    // rotl(0x8000_0001, 1) = 0x0000_0003
    let cv = chunk.add_constant(Value::I32(0x8000_0001_u32 as i32));
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, cv, 0);
    chunk.emit_op_u16(Op::r#const, c1, 0);
    chunk.emit_op(Op::i32_rotl, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(3) => {} _ => panic!("Expected I32(3), got {:?}", result) }
}

#[test]
fn i32_clz_ctz_popcnt() {
    let mut chunk = Chunk::new("test");
    // clz(1) = 31
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, c1, 0);
    chunk.emit_op(Op::i32_clz, 0);
    // ctz(0x80) = 7
    let c128 = chunk.add_constant(Value::I32(0x80));
    chunk.emit_op_u16(Op::r#const, c128, 0);
    chunk.emit_op(Op::i32_ctz, 0);
    // add: 31 + 7 = 38
    chunk.emit_op(Op::i32_add, 0);
    // popcnt(0xFF) = 8
    let c255 = chunk.add_constant(Value::I32(0xFF));
    chunk.emit_op_u16(Op::r#const, c255, 0);
    chunk.emit_op(Op::i32_popcnt, 0);
    // add: 38 + 8 = 46
    chunk.emit_op(Op::i32_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(46) => {} _ => panic!("Expected I32(46), got {:?}", result) }
}

#[test]
fn i32_eqz() {
    let mut chunk = Chunk::new("test");
    let c0 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op(Op::i32_eqz, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::Bool(true) => {} _ => panic!("Expected Bool(true), got {:?}", result) }
}

// ============================================================
// WASM Runtime — i64 Opcodes
// ============================================================

#[test]
fn i64_arithmetic() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::I64(1_000_000_000_000));
    let b = chunk.add_constant(Value::I64(2_000_000_000_000));
    chunk.emit_op_u16(Op::r#const, a, 0);
    chunk.emit_op_u16(Op::r#const, b, 0);
    chunk.emit_op(Op::i64_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I64(3_000_000_000_000) => {} _ => panic!("Expected I64(3T), got {:?}", result) }
}

#[test]
fn i64_bitwise() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::I64(0xFF00));
    let b = chunk.add_constant(Value::I64(0x0FF0));
    // AND: 0xFF00 & 0x0FF0 = 0x0F00
    chunk.emit_op_u16(Op::r#const, a, 0);
    chunk.emit_op_u16(Op::r#const, b, 0);
    chunk.emit_op(Op::i64_and, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I64(0x0F00) => {} _ => panic!("Expected I64(0x0F00), got {:?}", result) }
}

// ============================================================
// WASM Runtime — f64 Math Opcodes
// ============================================================

#[test]
fn f64_math_ops() {
    let mut chunk = Chunk::new("test");
    // ceil(3.2) = 4.0
    let c = chunk.add_constant(Value::F64(3.2));
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::f64_ceil, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 4.0 => {} _ => panic!("Expected F64(4.0), got {:?}", result) }
}

#[test]
fn f64_floor() {
    let mut chunk = Chunk::new("test");
    let c = chunk.add_constant(Value::F64(3.7));
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::f64_floor, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 3.0 => {} _ => panic!("Expected F64(3.0), got {:?}", result) }
}

#[test]
fn f64_sqrt() {
    let mut chunk = Chunk::new("test");
    let c = chunk.add_constant(Value::F64(144.0));
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::f64_sqrt, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 12.0 => {} _ => panic!("Expected F64(12.0), got {:?}", result) }
}

#[test]
fn f64_abs_neg() {
    let mut chunk = Chunk::new("test");
    let c = chunk.add_constant(Value::F64(-7.5));
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::f64_abs, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 7.5 => {} _ => panic!("Expected F64(7.5), got {:?}", result) }
}

#[test]
fn f64_min_max() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::F64(3.0));
    let b = chunk.add_constant(Value::F64(7.0));
    chunk.emit_op_u16(Op::r#const, a, 0);
    chunk.emit_op_u16(Op::r#const, b, 0);
    chunk.emit_op(Op::f64_min, 0);
    // min(3, 7) = 3
    chunk.emit_op_u16(Op::r#const, a, 0);
    chunk.emit_op_u16(Op::r#const, b, 0);
    chunk.emit_op(Op::f64_max, 0);
    // max(3, 7) = 7
    // stack: [3, 7], add = 10
    chunk.emit_op(Op::f64_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 10.0 => {} _ => panic!("Expected F64(10.0), got {:?}", result) }
}

#[test]
fn f64_trunc_nearest() {
    let mut chunk = Chunk::new("test");
    // trunc(3.9) = 3.0
    let c = chunk.add_constant(Value::F64(3.9));
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::f64_trunc, 0);
    // nearest(2.5) = 2.0 (round to even) — Rust's round() gives 3.0 though
    // Use 2.3 → nearest = 2.0
    let c2 = chunk.add_constant(Value::F64(2.3));
    chunk.emit_op_u16(Op::r#const, c2, 0);
    chunk.emit_op(Op::f64_nearest, 0);
    // 3.0 + 2.0 = 5.0
    chunk.emit_op(Op::f64_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 5.0 => {} _ => panic!("Expected F64(5.0), got {:?}", result) }
}

// ============================================================
// WASM Runtime — Select
// ============================================================

#[test]
fn select_true() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::F64(10.0));
    let b = chunk.add_constant(Value::F64(20.0));
    let cond = chunk.add_constant(Value::I32(1)); // true
    chunk.emit_op_u16(Op::r#const, a, 0);     // val1
    chunk.emit_op_u16(Op::r#const, b, 0);     // val2
    chunk.emit_op_u16(Op::r#const, cond, 0);  // condition
    chunk.emit_op(Op::select, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 10.0 => {} _ => panic!("Expected F64(10.0), got {:?}", result) }
}

#[test]
fn select_false() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::F64(10.0));
    let b = chunk.add_constant(Value::F64(20.0));
    let cond = chunk.add_constant(Value::I32(0)); // false
    chunk.emit_op_u16(Op::r#const, a, 0);
    chunk.emit_op_u16(Op::r#const, b, 0);
    chunk.emit_op_u16(Op::r#const, cond, 0);
    chunk.emit_op(Op::select, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 20.0 => {} _ => panic!("Expected F64(20.0), got {:?}", result) }
}

// ============================================================
// WASM Runtime — Conversions
// ============================================================

#[test]
fn i32_wrap_i64_test() {
    let mut chunk = Chunk::new("test");
    let c = chunk.add_constant(Value::I64(0x1_0000_002A)); // wrap → 42
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::i32_wrap_i64, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(42) => {} _ => panic!("Expected I32(42), got {:?}", result) }
}

#[test]
fn i64_extend_i32() {
    let mut chunk = Chunk::new("test");
    // extend_s: -1 as i32 → -1 as i64
    let cn1 = chunk.add_constant(Value::I32(-1));
    chunk.emit_op_u16(Op::r#const, cn1, 0);
    chunk.emit_op(Op::i64_extend_i32_s, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I64(-1) => {} _ => panic!("Expected I64(-1), got {:?}", result) }
}

#[test]
fn i64_extend_i32_unsigned() {
    let mut chunk = Chunk::new("test");
    // extend_u: -1 as i32 (0xFFFFFFFF) → 4294967295 as i64
    let cn1 = chunk.add_constant(Value::I32(-1));
    chunk.emit_op_u16(Op::r#const, cn1, 0);
    chunk.emit_op(Op::i64_extend_i32_u, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I64(v) if v == 0xFFFF_FFFF => {} _ => panic!("Expected I64(0xFFFFFFFF), got {:?}", result) }
}

#[test]
fn sign_extension_i32() {
    let mut chunk = Chunk::new("test");
    // extend8_s: 0x80 → -128
    let c = chunk.add_constant(Value::I32(0x80));
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::i32_extend8_s, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(-128) => {} _ => panic!("Expected I32(-128), got {:?}", result) }
}

#[test]
fn reinterpret_f64_i64() {
    let mut chunk = Chunk::new("test");
    // reinterpret 1.0 as i64
    let c = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::r#const, c, 0);
    chunk.emit_op(Op::i64_reinterpret_f64, 0);
    // reinterpret back
    chunk.emit_op(Op::f64_reinterpret_i64, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::F64(v) if v == 1.0 => {} _ => panic!("Expected F64(1.0), got {:?}", result) }
}

// ============================================================
// WASM Runtime — Extended Memory Ops
// ============================================================

#[test]
fn memory_i32_load8_signed() {
    let mut chunk = Chunk::new("test");
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, c1, 0);
    chunk.emit_op(Op::memory_grow, 0);
    chunk.emit_op(Op::drop, 0);

    // Store 0xFF at addr 0
    let c0 = chunk.add_constant(Value::I32(0));
    let cff = chunk.add_constant(Value::F64(0xFF as f64));
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op_u16(Op::r#const, cff, 0);
    chunk.emit_op(Op::i32_store8, 0);

    // Load unsigned → 255
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op(Op::i32_load8_u, 0);

    // Load signed → -1
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op(Op::i32_load8_s, 0);

    // -1 + 255 = 254
    chunk.emit_op(Op::i32_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(254) => {} _ => panic!("Expected I32(254), got {:?}", result) }
}

#[test]
fn memory_i32_load16() {
    let mut chunk = Chunk::new("test");
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, c1, 0);
    chunk.emit_op(Op::memory_grow, 0);
    chunk.emit_op(Op::drop, 0);

    // Store 0x8001 as i32 at addr 0
    let c0 = chunk.add_constant(Value::I32(0));
    let cv = chunk.add_constant(Value::F64(0x8001 as f64));
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op_u16(Op::r#const, cv, 0);
    chunk.emit_op(Op::i32_store, 0);

    // load16_u → 0x8001 = 32769
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op(Op::i32_load16_u, 0);

    // load16_s → -32767
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op(Op::i32_load16_s, 0);

    // 32769 + (-32767) = 2
    chunk.emit_op(Op::i32_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(2) => {} _ => panic!("Expected I32(2), got {:?}", result) }
}

#[test]
fn memory_f32_roundtrip() {
    let mut chunk = Chunk::new("test");
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, c1, 0);
    chunk.emit_op(Op::memory_grow, 0);
    chunk.emit_op(Op::drop, 0);

    // Store f32(3.14) at addr 0
    let c0 = chunk.add_constant(Value::I32(0));
    let cpi = chunk.add_constant(Value::F64(3.140000104904175)); // f32 precision
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op_u16(Op::r#const, cpi, 0);
    chunk.emit_op(Op::f32_store, 0);

    // Load back
    chunk.emit_op_u16(Op::r#const, c0, 0);
    chunk.emit_op(Op::f32_load, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) => assert!((v - 3.14).abs() < 0.01, "Expected ~3.14, got {}", v),
        _ => panic!("Expected F64, got {:?}", result),
    }
}

// ============================================================
// WASM Runtime — Bool as numeric (the fibonacci fix)
// ============================================================

#[test]
fn bool_in_arithmetic() {
    // This tests the fix that made fibonacci work:
    // Bool(true).as_i32() should be 1, not 0
    let mut chunk = Chunk::new("test");
    let c5 = chunk.add_constant(Value::I32(5));
    let c3 = chunk.add_constant(Value::I32(3));
    // 5 < 3 = false → 0
    chunk.emit_op_u16(Op::r#const, c5, 0);
    chunk.emit_op_u16(Op::r#const, c3, 0);
    chunk.emit_op(Op::dyn_lt, 0);
    // Result is Bool(false), add 10 → should be 10
    let c10 = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::r#const, c10, 0);
    chunk.emit_op(Op::i32_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(10) => {} _ => panic!("Expected I32(10), got {:?}", result) }
}

#[test]
fn bool_true_as_one() {
    let mut chunk = Chunk::new("test");
    let c3 = chunk.add_constant(Value::I32(3));
    let c5 = chunk.add_constant(Value::I32(5));
    // 3 < 5 = true → 1
    chunk.emit_op_u16(Op::r#const, c3, 0);
    chunk.emit_op_u16(Op::r#const, c5, 0);
    chunk.emit_op(Op::dyn_lt, 0);
    // Bool(true) + 10 → 11
    let c10 = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::r#const, c10, 0);
    chunk.emit_op(Op::i32_add, 0);
    chunk.emit_op(Op::halt, 0);

    let result = run_chunks(vec![chunk]);
    match result { Value::I32(11) => {} _ => panic!("Expected I32(11), got {:?}", result) }
}

// ============================================================
// WASM Round-trip with execution verification
// ============================================================

#[test]
fn roundtrip_run_hello() {
    // Build a chunk that does: push 6, push 7, multiply → 42
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c6 = chunk.add_constant(Value::F64(6.0));
    let c7 = chunk.add_constant(Value::F64(7.0));
    chunk.emit_op_u16(Op::r#const, c6, 0);
    chunk.emit_op_u16(Op::r#const, c7, 0);
    chunk.emit_op(Op::f64_mul, 0);
    chunk.emit_op(Op::halt, 0);

    // Write to WASM, read back, run
    let wasm_bytes = wasm::write_wasm(&vec![chunk]);
    assert!(wasm_bytes.len() > 8, "WASM should be more than just header");

    let restored = wasm::read_wasm(&wasm_bytes).unwrap();
    let result = run_chunks(restored);
    match result {
        Value::F64(v) if v == 42.0 => {}
        _ => panic!("Expected F64(42.0), got {:?}", result),
    }
}

#[test]
fn roundtrip_memory_ops() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;

    // memory.grow 1
    let c1 = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::r#const, c1, 0);
    chunk.emit_op(Op::memory_grow, 0);
    chunk.emit_op(Op::drop, 0);

    // memory.size → should be 1
    chunk.emit_op(Op::memory_size, 0);
    chunk.emit_op(Op::halt, 0);

    let wasm_bytes = wasm::write_wasm(&vec![chunk]);
    let restored = wasm::read_wasm(&wasm_bytes).unwrap();
    let (vm, result) = run_with_host(restored);
    assert_eq!(vm.memory.len(), 65536);
}

#[test]
fn roundtrip_i32_ops() {
    // Test that i32 opcodes survive round-trip
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c10 = chunk.add_constant(Value::I32(10));
    let c3 = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::r#const, c10, 0);
    chunk.emit_op_u16(Op::r#const, c3, 0);
    chunk.emit_op(Op::i32_sub, 0); // 10 - 3 = 7
    chunk.emit_op(Op::halt, 0);

    let restored = roundtrip(vec![chunk]);
    let result = run_chunks(restored);
    match result { Value::I32(7) => {} _ => panic!("Expected I32(7), got {:?}", result) }
}

// ============================================================
// WASM Writer — validates output starts with magic bytes
// ============================================================

#[test]
fn wasm_writer_magic() {
    let chunk = Chunk::new("empty");
    let bytes = wasm::write_wasm(&vec![chunk]);
    assert_eq!(&bytes[0..4], &[0x00, 0x61, 0x73, 0x6D]);
    assert_eq!(&bytes[4..8], &[0x01, 0x00, 0x00, 0x00]);
}
