use std::sync::Arc;
use vybe_runtime::value::*;
use vybe_runtime::*;
use vybe_platform_wasm as wasm;

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
    chunk.emit_op_u16(Op::CONST, ci, 0);
    chunk.emit_op(Op::HALT, 0);
    chunk.local_count = 1;

    let chunks = roundtrip(vec![chunk]);
    assert!(!chunks.is_empty());
}

#[test]
fn roundtrip_preserves_chunk_count() {
    let mut c1 = Chunk::new("main");
    c1.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c1.emit_op(Op::HALT, 0);

    let mut c2 = Chunk::new("helper");
    c2.arity = 1;
    c2.local_count = 2;
    c2.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c2.emit_op(Op::RETURN, 0);

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
    chunk.emit_op_u16(Op::CONST, c10, 0);
    chunk.emit_op_u16(Op::CONST, c3, 0);
    chunk.emit_op(Op::F64_DIV, 0);
    chunk.emit_op(Op::HALT, 0);

    // Run original
    let result1 = run_chunks(vec![chunk.clone()]);

    // Round-trip and run
    let restored = roundtrip(vec![chunk]);
    let result2 = run_chunks(restored);

    match (&result1, &result2) {
        (Value::F64(a), Value::F64(b)) => {
            assert!((a - b).abs() < 1e-10, "Results differ: {} vs {}", a, b);
        }
        _ => panic!("Expected F64 results") }
}

#[test]
fn roundtrip_i32_arithmetic() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c7 = chunk.add_constant(Value::I32(7));
    let c3 = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::CONST, c7, 0);
    chunk.emit_op_u16(Op::CONST, c3, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let restored = roundtrip(vec![chunk]);
    let result = run_chunks(restored);
    match result {
        Value::I32(10) => {}
        _ => panic!("Expected I32(10), got {:?}", result) }
}

#[test]
fn roundtrip_locals() {
    // set local 1 = 99, get local 1
    let mut chunk = Chunk::new("test");
    chunk.local_count = 3;
    let c99 = chunk.add_constant(Value::F64(99.0));
    chunk.emit_op_u16(Op::CONST, c99, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op(Op::HALT, 0);

    let restored = roundtrip(vec![chunk]);
    let result = run_chunks(restored);
    match result {
        Value::F64(v) if v == 99.0 => {}
        _ => panic!("Expected F64(99), got {:?}", result) }
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
    chunk.emit_op_u16(Op::CONST, c17, 0);
    chunk.emit_op_u16(Op::CONST, c5, 0);
    chunk.emit_op(Op::I32_DIV_S, 0);

    // 17 % 5 = 2
    chunk.emit_op_u16(Op::CONST, c17, 0);
    chunk.emit_op_u16(Op::CONST, c5, 0);
    chunk.emit_op(Op::I32_REM_S, 0);

    // Stack: [3, 2] → add = 5
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(5) => {}
        _ => panic!("Expected I32(5), got {:?}", result) }
}

#[test]
fn i32_div_by_zero() {
    // WASM spec: i32.div_s with zero divisor traps ("integer divide by zero").
    let mut chunk = Chunk::new("test");
    let c10 = chunk.add_constant(Value::I32(10));
    let c0 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, c10, 0);
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::I32_DIV_S, 0);
    chunk.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    let err = vm.run(vec![chunk]).expect_err("expected trap");
    assert!(
        err.message.contains("divide by zero"),
        "got: {}",
        err.message
    );
}

#[test]
fn i32_rotate() {
    let mut chunk = Chunk::new("test");
    // rotl(0x8000_0001, 1) = 0x0000_0003
    let cv = chunk.add_constant(Value::I32(0x8000_0001_u32 as i32));
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, cv, 0);
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::I32_ROTL, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(3) => {}
        _ => panic!("Expected I32(3), got {:?}", result) }
}

#[test]
fn i32_clz_ctz_popcnt() {
    let mut chunk = Chunk::new("test");
    // clz(1) = 31
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::I32_CLZ, 0);
    // ctz(0x80) = 7
    let c128 = chunk.add_constant(Value::I32(0x80));
    chunk.emit_op_u16(Op::CONST, c128, 0);
    chunk.emit_op(Op::I32_CTZ, 0);
    // add: 31 + 7 = 38
    chunk.emit_op(Op::I32_ADD, 0);
    // popcnt(0xFF) = 8
    let c255 = chunk.add_constant(Value::I32(0xFF));
    chunk.emit_op_u16(Op::CONST, c255, 0);
    chunk.emit_op(Op::I32_POPCNT, 0);
    // add: 38 + 8 = 46
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(46) => {}
        _ => panic!("Expected I32(46), got {:?}", result) }
}

#[test]
fn i32_eqz() {
    let mut chunk = Chunk::new("test");
    let c0 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::I32_EQZ, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    // i32.eqz returns i32(1) per WASM spec §3.3.4.1 (not bool)
    assert_eq!(result.as_i32(), 1);
}

// ============================================================
// WASM Runtime — i64 Opcodes
// ============================================================

#[test]
fn i64_arithmetic() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::I64(1_000_000_000_000));
    let b = chunk.add_constant(Value::I64(2_000_000_000_000));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I64_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I64(3_000_000_000_000) => {}
        _ => panic!("Expected I64(3T), got {:?}", result) }
}

#[test]
fn i64_bitwise() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::I64(0xFF00));
    let b = chunk.add_constant(Value::I64(0x0FF0));
    // AND: 0xFF00 & 0x0FF0 = 0x0F00
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I64_AND, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I64(0x0F00) => {}
        _ => panic!("Expected I64(0x0F00), got {:?}", result) }
}

// ============================================================
// WASM Runtime — f64 Math Opcodes
// ============================================================

#[test]
fn f64_math_ops() {
    let mut chunk = Chunk::new("test");
    // ceil(3.2) = 4.0
    let c = chunk.add_constant(Value::F64(3.2));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::F64_CEIL, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 4.0 => {}
        _ => panic!("Expected F64(4.0), got {:?}", result) }
}

#[test]
fn f64_floor() {
    let mut chunk = Chunk::new("test");
    let c = chunk.add_constant(Value::F64(3.7));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::F64_FLOOR, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 3.0 => {}
        _ => panic!("Expected F64(3.0), got {:?}", result) }
}

#[test]
fn f64_sqrt() {
    let mut chunk = Chunk::new("test");
    let c = chunk.add_constant(Value::F64(144.0));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::F64_SQRT, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 12.0 => {}
        _ => panic!("Expected F64(12.0), got {:?}", result) }
}

#[test]
fn f64_abs_neg() {
    let mut chunk = Chunk::new("test");
    let c = chunk.add_constant(Value::F64(-7.5));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::F64_ABS, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 7.5 => {}
        _ => panic!("Expected F64(7.5), got {:?}", result) }
}

#[test]
fn f64_min_max() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::F64(3.0));
    let b = chunk.add_constant(Value::F64(7.0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::F64_MIN, 0);
    // min(3, 7) = 3
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::F64_MAX, 0);
    // max(3, 7) = 7
    // stack: [3, 7], add = 10
    chunk.emit_op(Op::F64_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 10.0 => {}
        _ => panic!("Expected F64(10.0), got {:?}", result) }
}

#[test]
fn f64_trunc_nearest() {
    let mut chunk = Chunk::new("test");
    // trunc(3.9) = 3.0
    let c = chunk.add_constant(Value::F64(3.9));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::F64_TRUNC, 0);
    // nearest(2.5) = 2.0 (round to even) — Rust's round() gives 3.0 though
    // Use 2.3 → nearest = 2.0
    let c2 = chunk.add_constant(Value::F64(2.3));
    chunk.emit_op_u16(Op::CONST, c2, 0);
    chunk.emit_op(Op::F64_NEAREST, 0);
    // 3.0 + 2.0 = 5.0
    chunk.emit_op(Op::F64_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 5.0 => {}
        _ => panic!("Expected F64(5.0), got {:?}", result) }
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
    chunk.emit_op_u16(Op::CONST, a, 0); // val1
    chunk.emit_op_u16(Op::CONST, b, 0); // val2
    chunk.emit_op_u16(Op::CONST, cond, 0); // condition
    chunk.emit_op(Op::SELECT, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 10.0 => {}
        _ => panic!("Expected F64(10.0), got {:?}", result) }
}

#[test]
fn select_false() {
    let mut chunk = Chunk::new("test");
    let a = chunk.add_constant(Value::F64(10.0));
    let b = chunk.add_constant(Value::F64(20.0));
    let cond = chunk.add_constant(Value::I32(0)); // false
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, cond, 0);
    chunk.emit_op(Op::SELECT, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 20.0 => {}
        _ => panic!("Expected F64(20.0), got {:?}", result) }
}

// ============================================================
// WASM Runtime — Conversions
// ============================================================

#[test]
fn i32_wrap_i64_test() {
    let mut chunk = Chunk::new("test");
    let c = chunk.add_constant(Value::I64(0x1_0000_002A)); // wrap → 42
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::I32_WRAP_I64, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(42) => {}
        _ => panic!("Expected I32(42), got {:?}", result) }
}

#[test]
fn i64_extend_i32() {
    let mut chunk = Chunk::new("test");
    // extend_s: -1 as i32 → -1 as i64
    let cn1 = chunk.add_constant(Value::I32(-1));
    chunk.emit_op_u16(Op::CONST, cn1, 0);
    chunk.emit_op(Op::I64_EXTEND_I32_S, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I64(-1) => {}
        _ => panic!("Expected I64(-1), got {:?}", result) }
}

#[test]
fn i64_extend_i32_unsigned() {
    let mut chunk = Chunk::new("test");
    // extend_u: -1 as i32 (0xFFFFFFFF) → 4294967295 as i64
    let cn1 = chunk.add_constant(Value::I32(-1));
    chunk.emit_op_u16(Op::CONST, cn1, 0);
    chunk.emit_op(Op::I64_EXTEND_I32_U, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I64(v) if v == 0xFFFF_FFFF => {}
        _ => panic!("Expected I64(0xFFFFFFFF), got {:?}", result) }
}

#[test]
fn sign_extension_i32() {
    let mut chunk = Chunk::new("test");
    // extend8_s: 0x80 → -128
    let c = chunk.add_constant(Value::I32(0x80));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::I32_EXTEND8_S, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(-128) => {}
        _ => panic!("Expected I32(-128), got {:?}", result) }
}

#[test]
fn reinterpret_f64_i64() {
    let mut chunk = Chunk::new("test");
    // reinterpret 1.0 as i64
    let c = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::I64_REINTERPRET_F64, 0);
    // reinterpret back
    chunk.emit_op(Op::F64_REINTERPRET_I64, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 1.0 => {}
        _ => panic!("Expected F64(1.0), got {:?}", result) }
}

// ============================================================
// WASM Runtime — Extended Memory Ops
// ============================================================

#[test]
fn memory_i32_load8_signed() {
    let mut chunk = Chunk::new("test");
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);

    // Store 0xFF at addr 0
    let c0 = chunk.add_constant(Value::I32(0));
    let cff = chunk.add_constant(Value::F64(0xFF as f64));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, cff, 0);
    chunk.emit_op(Op::I32_STORE8, 0);

    // Load unsigned → 255
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);

    // Load signed → -1
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::I32_LOAD8_S, 0);

    // -1 + 255 = 254
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(254) => {}
        _ => panic!("Expected I32(254), got {:?}", result) }
}

#[test]
fn memory_i32_load16() {
    let mut chunk = Chunk::new("test");
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);

    // Store 0x8001 as i32 at addr 0
    let c0 = chunk.add_constant(Value::I32(0));
    let cv = chunk.add_constant(Value::F64(0x8001 as f64));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, cv, 0);
    chunk.emit_op(Op::I32_STORE, 0);

    // load16_u → 0x8001 = 32769
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::I32_LOAD16_U, 0);

    // load16_s → -32767
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::I32_LOAD16_S, 0);

    // 32769 + (-32767) = 2
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(2) => {}
        _ => panic!("Expected I32(2), got {:?}", result) }
}

#[test]
fn memory_f32_roundtrip() {
    let mut chunk = Chunk::new("test");
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);

    // Store f32(3.14) at addr 0
    let c0 = chunk.add_constant(Value::I32(0));
    let cpi = chunk.add_constant(Value::F64(3.140000104904175)); // f32 precision
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, cpi, 0);
    chunk.emit_op(Op::F32_STORE, 0);

    // Load back
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::F32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        // f32.load loads an f32 (spec) → Value::F32.
        Value::F32(v) => assert!((v as f64 - 3.14).abs() < 0.01, "Expected ~3.14, got {}", v),
        _ => panic!("Expected F32, got {:?}", result) }
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
    chunk.emit_op_u16(Op::CONST, c5, 0);
    chunk.emit_op_u16(Op::CONST, c3, 0);
    chunk.emit_op(Op::I32_LT_S, 0);
    // Result is Bool(false), add 10 → should be 10
    let c10 = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::CONST, c10, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(10) => {}
        _ => panic!("Expected I32(10), got {:?}", result) }
}

#[test]
fn bool_true_as_one() {
    let mut chunk = Chunk::new("test");
    let c3 = chunk.add_constant(Value::I32(3));
    let c5 = chunk.add_constant(Value::I32(5));
    // 3 < 5 = true → 1
    chunk.emit_op_u16(Op::CONST, c3, 0);
    chunk.emit_op_u16(Op::CONST, c5, 0);
    chunk.emit_op(Op::I32_LT_S, 0);
    // Bool(true) + 10 → 11
    let c10 = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::CONST, c10, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(11) => {}
        _ => panic!("Expected I32(11), got {:?}", result) }
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
    chunk.emit_op_u16(Op::CONST, c6, 0);
    chunk.emit_op_u16(Op::CONST, c7, 0);
    chunk.emit_op(Op::F64_MUL, 0);
    chunk.emit_op(Op::HALT, 0);

    // Write to WASM, read back, run
    let wasm_bytes = wasm::write_wasm(&vec![chunk]);
    assert!(wasm_bytes.len() > 8, "WASM should be more than just header");

    let restored = wasm::read_wasm(&wasm_bytes).unwrap();
    let result = run_chunks(restored);
    match result {
        Value::F64(v) if v == 42.0 => {}
        _ => panic!("Expected F64(42.0), got {:?}", result) }
}

#[test]
fn roundtrip_memory_ops() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;

    // memory.grow 1
    let c1 = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);

    // memory.size → should be 1
    chunk.emit_op(Op::MEMORY_SIZE, 0);
    chunk.emit_op(Op::HALT, 0);

    let wasm_bytes = wasm::write_wasm(&vec![chunk]);
    let restored = wasm::read_wasm(&wasm_bytes).unwrap();
    let (vm, _result) = run_with_host(restored);
    assert_eq!(vm.memory.len(), 65536);
}

#[test]
fn roundtrip_i32_ops() {
    // Test that i32 opcodes survive round-trip
    let mut chunk = Chunk::new("test");
    chunk.local_count = 1;
    let c10 = chunk.add_constant(Value::I32(10));
    let c3 = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::CONST, c10, 0);
    chunk.emit_op_u16(Op::CONST, c3, 0);
    chunk.emit_op(Op::I32_SUB, 0); // 10 - 3 = 7
    chunk.emit_op(Op::HALT, 0);

    let restored = roundtrip(vec![chunk]);
    let result = run_chunks(restored);
    match result {
        Value::I32(7) => {}
        _ => panic!("Expected I32(7), got {:?}", result) }
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

// ============================================================
// JS String Builtins
// ============================================================

// ============================================================
// Array Builtins
// ============================================================

#[test]
fn array_length_op() {
    let mut chunk = Chunk::new("test");
    let c1 = chunk.add_constant(Value::I32(10));
    let c2 = chunk.add_constant(Value::I32(20));
    let c3 = chunk.add_constant(Value::I32(30));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op_u16(Op::CONST, c2, 0);
    chunk.emit_op_u16(Op::CONST, c3, 0);
    chunk.emit_array_new_fixed(0, 3, 0);
    chunk.emit_op(Op::ARRAY_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(3) => {}
        _ => panic!("Expected I32(3), got {:?}", result) }
}

#[test]
fn array_new_default_op() {
    // Spec: `array.new_default $typeidx` takes a u16 typeidx immediate.
    let mut chunk = Chunk::new("test");
    let c5 = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, c5, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, 0);
    chunk.emit_op(Op::ARRAY_LENGTH, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(5) => {}
        _ => panic!("Expected I32(5), got {:?}", result) }
}

// ============================================================
// SIMD (i32x4)
// ============================================================

#[test]
fn simd_i32x4_add() {
    let mut chunk = Chunk::new("test");
    // splat 10 → [10, 10, 10, 10]
    let c10 = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::CONST, c10, 0);
    chunk.emit_op(Op::I32X4_SPLAT, 0);
    // splat 5 → [5, 5, 5, 5]
    let c5 = chunk.add_constant(Value::I32(5));
    chunk.emit_op_u16(Op::CONST, c5, 0);
    chunk.emit_op(Op::I32X4_SPLAT, 0);
    // add → [15, 15, 15, 15]
    chunk.emit_op(Op::I32X4_ADD, 0);
    // extract lane 0 → 15
    chunk.emit_op_u8(Op::I32X4_EXTRACT_LANE, 0, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(15) => {}
        _ => panic!("Expected I32(15), got {:?}", result) }
}

#[test]
fn simd_f64x2_mul() {
    let mut chunk = Chunk::new("test");
    let c3 = chunk.add_constant(Value::F64(3.0));
    chunk.emit_op_u16(Op::CONST, c3, 0);
    chunk.emit_op(Op::F64X2_SPLAT, 0);
    let c7 = chunk.add_constant(Value::F64(7.0));
    chunk.emit_op_u16(Op::CONST, c7, 0);
    chunk.emit_op(Op::F64X2_SPLAT, 0);
    chunk.emit_op(Op::F64X2_MUL, 0);
    chunk.emit_op_u8(Op::F64X2_EXTRACT_LANE, 0, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::F64(v) if v == 21.0 => {}
        _ => panic!("Expected F64(21.0), got {:?}", result) }
}

#[test]
fn simd_v128_bitwise() {
    let mut chunk = Chunk::new("test");
    // [0xFF, 0xFF, ...] AND [0x0F, 0x0F, ...] = [0x0F, 0x0F, ...]
    let c_ff = chunk.add_constant(Value::I32(0xFF));
    chunk.emit_op_u16(Op::CONST, c_ff, 0);
    chunk.emit_op(Op::I8X16_SPLAT, 0);
    let c_0f = chunk.add_constant(Value::I32(0x0F));
    chunk.emit_op_u16(Op::CONST, c_0f, 0);
    chunk.emit_op(Op::I8X16_SPLAT, 0);
    chunk.emit_op(Op::V128_AND, 0);
    chunk.emit_op_u8(Op::I8X16_EXTRACT_LANE_U, 0, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(0x0F) => {}
        _ => panic!("Expected I32(0x0F), got {:?}", result) }
}

// ============================================================
// Atomics (single-threaded correctness)
// ============================================================

fn emit_atomic(c: &mut Chunk, op: Op) {
    c.emit_op(op, 0);
    c.emit(0, 0);
    c.emit(0, 0);
}

#[test]
fn atomic_rmw_add() {
    let mut chunk = Chunk::new("test");
    // Grow memory
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);
    // Store 100 at addr 0
    let c0 = chunk.add_constant(Value::I32(0));
    let c100 = chunk.add_constant(Value::I32(100));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, c100, 0);
    emit_atomic(&mut chunk, Op::I32_ATOMIC_STORE);
    // atomic_rmw_add(0, 42) → old=100, new=142
    chunk.emit_op_u16(Op::CONST, c0, 0);
    let c42 = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, c42, 0);
    emit_atomic(&mut chunk, Op::I32_ATOMIC_RMW_ADD);
    // Returns old value (100)
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(100) => {}
        _ => panic!("Expected I32(100), got {:?}", result) }
}

#[test]
fn atomic_cmpxchg() {
    let mut chunk = Chunk::new("test");
    let c1 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);
    // Store 50 at addr 0
    let c0 = chunk.add_constant(Value::I32(0));
    let c50 = chunk.add_constant(Value::I32(50));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, c50, 0);
    emit_atomic(&mut chunk, Op::I32_ATOMIC_STORE);
    // cmpxchg(addr=0, expected=50, replacement=99) → old=50, swap happens
    let c99 = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, c50, 0);
    chunk.emit_op_u16(Op::CONST, c99, 0);
    emit_atomic(&mut chunk, Op::I32_ATOMIC_RMW_CMPXCHG);
    chunk.emit_op(Op::DROP, 0); // drop old (50)
    // Load → should be 99
    chunk.emit_op_u16(Op::CONST, c0, 0);
    emit_atomic(&mut chunk, Op::I32_ATOMIC_LOAD);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(99) => {}
        _ => panic!("Expected I32(99), got {:?}", result) }
}

// ============================================================
// i31ref (tagged small integers)
// ============================================================

#[test]
fn i31ref_roundtrip() {
    let mut chunk = Chunk::new("test");
    let c42 = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, c42, 0);
    chunk.emit_op(Op::I31_NEW, 0);
    chunk.emit_op(Op::I31_GET_S, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(42) => {}
        _ => panic!("Expected I32(42), got {:?}", result) }
}

#[test]
fn i31ref_negative() {
    let mut chunk = Chunk::new("test");
    // -1 & 0x7FFF_FFFF = 0x7FFF_FFFF, sign extend from bit 30 → -1
    let cn1 = chunk.add_constant(Value::I32(-1));
    chunk.emit_op_u16(Op::CONST, cn1, 0);
    chunk.emit_op(Op::I31_NEW, 0);
    chunk.emit_op(Op::I31_GET_S, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(-1) => {}
        _ => panic!("Expected I32(-1), got {:?}", result) }
}

// ============================================================
// ref_cast / br_on_cast
// ============================================================

#[test]
fn ref_cast_success() {
    let mut chunk = Chunk::new("test");
    let cs = chunk.add_constant(Value::String(Arc::from("hello")));
    chunk.emit_op_u16(Op::CONST, cs, 0);
    // Every value in this ABI is an external reference, so `ref.cast extern`
    // is the cast that must succeed.
    chunk.emit_ref_type_op(
        Op::REF_CAST,
        vybe_runtime::opcode::heaptype::HeapType::Abstract(
            vybe_runtime::opcode::heaptype::HT_EXTERN,
        ),
        0,
    );
    // Should not trap — the cast value stays on the stack unchanged.
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::String(s) if s.as_ref() == "hello" => {}
        _ => panic!("Expected String(\"hello\"), got {:?}", result) }
}

// ============================================================
// call_ref (typed function references)
// ============================================================

#[test]
fn call_ref_basic() {
    // Chunk 0: script — create closure, call via call_ref
    let mut script = Chunk::new("<script>");
    // Chunk 1: double function (x → x * 2)
    let mut double_chunk = Chunk::new("double");
    double_chunk.arity = 1;
    double_chunk.local_count = 1;
    double_chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let c2 = double_chunk.add_constant(Value::F64(2.0));
    double_chunk.emit_op_u16(Op::CONST, c2, 0);
    double_chunk.emit_op(Op::F64_MUL, 0);
    double_chunk.emit_op(Op::RETURN, 0);

    // Script: ref_func 1, push arg 21, call_ref 1
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0); // 0 upvalues
    let c21 = script.add_constant(Value::F64(21.0));
    script.emit_op_u16(Op::CONST, c21, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![script, double_chunk]);
    match result {
        Value::F64(v) if v == 42.0 => {}
        _ => panic!("Expected F64(42.0), got {:?}", result) }
}

#[test]
fn call_ref_preserves_array_argument() {
    let mut script = Chunk::new("<script>");
    let mut first_chunk = Chunk::new("first");
    first_chunk.arity = 1;
    first_chunk.local_count = 1;
    first_chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    first_chunk.emit_i32_const(0, 0);
    first_chunk.emit_op(Op::ARRAY_GET, 0);
    first_chunk.emit_op(Op::RETURN, 0);

    let array = Value::Object(Arc::new(std::sync::Mutex::new(Object::new_array(vec![
        Value::F64(50.0),
    ]))));
    let c_array = script.add_constant(array);
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::CONST, c_array, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![script, first_chunk]);
    match result {
        Value::F64(v) if v == 50.0 => {}
        _ => panic!("Expected F64(50.0), got {:?}", result) }
}

// ============================================================
// Memory64
// ============================================================

fn emit_leb_u64(out: &mut Chunk, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.emit(byte, 0);
        if value == 0 {
            break;
        }
    }
}

#[test]
fn memory64_grow_and_load() {
    fn emit_memarg64(out: &mut Chunk, align: u32, offset: u64, memidx: u32) {
        let encoded_align = if memidx == 0 { align } else { align | 0x40 };
        out.emit_leb_u32(encoded_align, 0);
        let mut value = offset;
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            out.emit(byte, 0);
            if value == 0 {
                break;
            }
        }
        if memidx != 0 {
            out.emit_leb_u32(memidx, 0);
        }
    }

    let mut chunk = Chunk::new("test");
    // Grow with i64
    let c1 = chunk.add_constant(Value::I64(1));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);
    // Store 42 at i64 addr 0
    let c0 = chunk.add_constant(Value::I64(0));
    let c42 = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, c42, 0);
    chunk.emit_op(Op::I32_STORE, 0);
    emit_memarg64(&mut chunk, 2, 0, 0);
    // Load back
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    emit_memarg64(&mut chunk, 2, 0, 0);
    chunk.emit_op(Op::HALT, 0);
    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(42) => {}
        _ => panic!("Expected I32(42), got {:?}", result) }
}

#[test]
fn memory64_load_store_apply_memarg_offset() {
    let mut chunk = Chunk::new("test");
    let pages = chunk.add_constant(Value::I64(1));
    let base = chunk.add_constant(Value::I64(4));
    let value = chunk.add_constant(Value::I32(99));

    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_op_u16(Op::CONST, base, 0);
    chunk.emit_op_u16(Op::CONST, value, 0);
    chunk.emit_op(Op::I32_STORE, 0);
    chunk.emit_leb_u32(2, 0); // align
    emit_leb_u64(&mut chunk, 8); // offset: effective address is 12

    chunk.emit_op_u16(Op::CONST, base, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_leb_u32(2, 0);
    emit_leb_u64(&mut chunk, 8);
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    assert_eq!(result.as_i32(), 99);
}

// ============================================================
// JSPI (JavaScript Promise Integration)
// ============================================================

fn make_promise(id: u64, state: &str, value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Promise")));
    obj.properties.insert("__id".into(), Value::F64(id as f64));
    obj.properties
        .insert("__state".into(), Value::String(Arc::from(state)));
    obj.properties.insert("__value".into(), value);
    Value::Object(Arc::new(std::sync::Mutex::new(obj)))
}

#[test]
fn jspi_resolved_promise_returns_immediately() {
    // Host function returns an already-resolved promise.
    // JSPI should extract the value without suspending.
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "fetch_sync",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            make_promise(1, "fulfilled", Value::String(Arc::from("data from server")))
        }),
    );

    let mut chunk = Chunk::new("<test>");
    let idx = chunk.add_import("test", "fetch_sync");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    chunk.emit(0, 0); // 0 args

    // The result is a fulfilled promise — JSPI should NOT suspend.
    // It should push the promise object (call_import doesn't auto-unwrap fulfilled).
    // To unwrap, we use promise_suspend opcode.
    {
        let aw = chunk.add_import("jspi", "await");
        chunk.emit_op_u16(Op::CALL_IMPORT, aw, 0);
        chunk.emit(1, 0);
    }
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    match &result {
        Value::String(s) if s.as_ref() == "data from server" => {}
        _ => panic!("Expected 'data from server', got {:?}", result) }
}

#[test]
fn jspi_non_promise_passes_through() {
    // Host function returns a plain value (not a promise).
    // JSPI should pass it through unchanged.
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "compute",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| Value::I32(42)),
    );

    let mut chunk = Chunk::new("<test>");
    let idx = chunk.add_import("test", "compute");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    chunk.emit(0, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    match result {
        Value::I32(42) => {}
        _ => panic!("Expected I32(42), got {:?}", result) }
}

#[test]
fn jspi_pending_promise_suspends() {
    // Host function returns a pending promise.
    // JSPI should suspend the fiber and return a special error.
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "slow_fetch",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            make_promise(99, "pending", Value::Null)
        }),
    );

    let mut chunk = Chunk::new("<test>");
    let idx = chunk.add_import("test", "slow_fetch");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    chunk.emit(0, 0);
    // `await slow_fetch()` — the JSPI suspend point. A non-suspending import
    // returning a pending promise does NOT itself suspend; suspension happens
    // at the explicit `await` (the `jspi.await` suspending import).
    let aw = chunk.add_import("jspi", "await");
    chunk.emit_op_u16(Op::CALL_IMPORT, aw, 0);
    chunk.emit(1, 0);
    // This should never reach — the await suspends via JSPI
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]);
    // Should get a JSPI suspension error
    match result {
        Err(e) => {
            let msg = format!("{}", e);
            assert!(
                msg.contains("__jspi__:99"),
                "Expected JSPI suspension, got: {}",
                msg
            );
        }
        Ok(v) => panic!("Expected JSPI suspension, got Ok({:?})", v) }
}

#[test]
fn jspi_suspend_then_resume() {
    // Full JSPI cycle: host returns pending → suspend → resolve → resume.
    let mut vm = VM::new();
    let output: Arc<std::sync::Mutex<Vec<String>>> = Arc::new(std::sync::Mutex::new(Vec::new()));
    let out = output.clone();

    vm.register_host_fn(
        "test",
        "log",
        Box::new(
            move |_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
                out.lock()
                    .unwrap()
                    .push(format!("{}", args.first().unwrap_or(&Value::Null)));
                Value::Null
            },
        ),
    );
    vm.register_host_fn(
        "test",
        "async_load",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            make_promise(42, "pending", Value::Null)
        }),
    );

    let mut chunk = Chunk::new("<test>");

    // Step 1: log "before"
    let msg1 = chunk.add_constant(Value::String(Arc::from("before")));
    chunk.emit_op_u16(Op::CONST, msg1, 0);
    let log_idx = chunk.add_import("test", "log");
    chunk.emit_op_u16(Op::CALL_IMPORT, log_idx, 0);
    chunk.emit(1, 0);
    chunk.emit_op(Op::DROP, 0);

    // Step 2: `await async_load()` — the JSPI suspend point.
    let load_idx = chunk.add_import("test", "async_load");
    chunk.emit_op_u16(Op::CALL_IMPORT, load_idx, 0);
    chunk.emit(0, 0);
    let aw = chunk.add_import("jspi", "await");
    chunk.emit_op_u16(Op::CALL_IMPORT, aw, 0);
    chunk.emit(1, 0);

    // Step 3: log the result (only reached after resume)
    chunk.emit_op_u16(Op::CALL_IMPORT, log_idx, 0);
    chunk.emit(1, 0);
    chunk.emit_op(Op::DROP, 0);

    // Step 4: log "after"
    let msg2 = chunk.add_constant(Value::String(Arc::from("after")));
    chunk.emit_op_u16(Op::CONST, msg2, 0);
    chunk.emit_op_u16(Op::CALL_IMPORT, log_idx, 0);
    chunk.emit(1, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_op(Op::HALT, 0);

    // Run — should suspend at async_load
    let result = vm.run(vec![chunk]);
    assert!(result.is_err());
    assert_eq!(*output.lock().unwrap(), vec!["before"]);
    assert!(vm.has_pending_jspi());

    // Resolve the promise — this resumes execution
    vm.jspi_resolve(42, Value::String(Arc::from("loaded data")))
        .unwrap();

    // Now "after" should have been logged
    assert_eq!(
        *output.lock().unwrap(),
        vec!["before", "loaded data", "after"]
    );
    assert!(!vm.has_pending_jspi());
}

#[test]
fn jspi_promise_suspend_opcode() {
    // Test the promise_suspend opcode with a fulfilled promise
    let mut chunk = Chunk::new("<test>");
    // Create a fulfilled promise manually via constants
    let type_k = chunk.add_constant(Value::String(Arc::from("__type")));
    let type_v = chunk.add_constant(Value::String(Arc::from("Promise")));
    let state_k = chunk.add_constant(Value::String(Arc::from("__state")));
    let state_v = chunk.add_constant(Value::String(Arc::from("fulfilled")));
    let value_k = chunk.add_constant(Value::String(Arc::from("__value")));
    let value_v = chunk.add_constant(Value::I32(99));

    // Build object: {__type: "Promise", __state: "fulfilled", __value: 99}
    chunk.emit_op_u16(Op::CONST, type_k, 0);
    chunk.emit_op_u16(Op::CONST, type_v, 0);
    chunk.emit_op_u16(Op::CONST, state_k, 0);
    chunk.emit_op_u16(Op::CONST, state_v, 0);
    chunk.emit_op_u16(Op::CONST, value_k, 0);
    chunk.emit_op_u16(Op::CONST, value_v, 0);
    chunk.emit_struct_new(0, 3, 0);

    // promise_suspend should extract the value
    {
        let aw = chunk.add_import("jspi", "await");
        chunk.emit_op_u16(Op::CALL_IMPORT, aw, 0);
        chunk.emit(1, 0);
    }
    chunk.emit_op(Op::HALT, 0);

    let result = run_chunks(vec![chunk]);
    match result {
        Value::I32(99) => {}
        _ => panic!("Expected I32(99), got {:?}", result) }
}
