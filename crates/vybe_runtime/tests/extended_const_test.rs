use std::sync::Arc;
use vybe_runtime::chunk::ConstExpr;
/// Tests for Extended Const Expressions, Typed Continuations, and String References.
use vybe_runtime::{Chunk, Op, VM, Value};

// ── Extended Const Expressions ──────────────────────────────

#[test]
fn global_init_literal() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.add_global_init("PI", ConstExpr::Value(Value::F64(3.14159)));

    // Read the global
    let name = chunk.add_constant(Value::String(Arc::from("PI")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!((result.as_f64() - 3.14159).abs() < 1e-10);
}

#[test]
fn global_init_add() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    // BASE = 100, OFFSET = BASE + 50
    chunk.add_global_init("BASE", ConstExpr::Value(Value::I32(100)));
    chunk.add_global_init(
        "OFFSET",
        ConstExpr::Add(
            Box::new(ConstExpr::GlobalGet("BASE".into())),
            Box::new(ConstExpr::Value(Value::I32(50))),
        ),
    );

    let name = chunk.add_constant(Value::String(Arc::from("OFFSET")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 150);
}

#[test]
fn global_init_mul() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    // SIZE = 4, TABLE = SIZE * 256
    chunk.add_global_init("SIZE", ConstExpr::Value(Value::I32(4)));
    chunk.add_global_init(
        "TABLE",
        ConstExpr::Mul(
            Box::new(ConstExpr::GlobalGet("SIZE".into())),
            Box::new(ConstExpr::Value(Value::I32(256))),
        ),
    );

    let name = chunk.add_constant(Value::String(Arc::from("TABLE")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 1024);
}

#[test]
fn global_init_chain() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    // A = 10, B = A + 20, C = B + A → 40
    chunk.add_global_init("A", ConstExpr::Value(Value::I32(10)));
    chunk.add_global_init(
        "B",
        ConstExpr::Add(
            Box::new(ConstExpr::GlobalGet("A".into())),
            Box::new(ConstExpr::Value(Value::I32(20))),
        ),
    );
    chunk.add_global_init(
        "C",
        ConstExpr::Add(
            Box::new(ConstExpr::GlobalGet("B".into())),
            Box::new(ConstExpr::GlobalGet("A".into())),
        ),
    );

    let name = chunk.add_constant(Value::String(Arc::from("C")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 40);
}

#[test]
fn global_init_missing_global_get_yields_null() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.add_global_init("MISSING_REF", ConstExpr::GlobalGet("DOES_NOT_EXIST".into()));

    let name = chunk.add_constant(Value::String(Arc::from("MISSING_REF")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Null));
}

#[test]
fn global_init_i64_arithmetic_wraps() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.add_global_init(
        "I64_WRAP",
        ConstExpr::Add(
            Box::new(ConstExpr::Value(Value::I64(i64::MAX))),
            Box::new(ConstExpr::Value(Value::I64(1))),
        ),
    );

    let name = chunk.add_constant(Value::String(Arc::from("I64_WRAP")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i64(), i64::MIN);
}

#[test]
fn global_init_f64_arithmetic_preserves_float_type() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.add_global_init(
        "F64_PRODUCT",
        ConstExpr::Mul(
            Box::new(ConstExpr::Value(Value::F64(1.5))),
            Box::new(ConstExpr::Value(Value::F64(2.0))),
        ),
    );

    let name = chunk.add_constant(Value::String(Arc::from("F64_PRODUCT")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 3.0);
}

// ── Typed Continuations ─────────────────────────────────────

#[test]
fn continuation_tag_in_chunk() {
    let mut chunk = Chunk::new("<script>");
    let idx0 = chunk.add_continuation_tag("gen_int", "i32", "i32");
    let idx1 = chunk.add_continuation_tag("gen_str", "string", "string");

    assert_eq!(idx0, 0);
    assert_eq!(idx1, 1);
    assert_eq!(chunk.continuation_tags.len(), 2);
    assert_eq!(chunk.continuation_tags[0].name, "gen_int");
    assert_eq!(chunk.continuation_tags[1].yield_type, "string");
}

// ── String References ───────────────────────────────────────
