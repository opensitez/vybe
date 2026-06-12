use std::sync::Arc;
use vybe_bytecode::chunk::ConstExpr;
/// Tests for Extended Const Expressions, Typed Continuations, and String References.
use vybe_bytecode::{Chunk, Op, VM, Value};

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
    chunk.emit_op(Op::HALT, 0);

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
    chunk.emit_op(Op::HALT, 0);

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
    chunk.emit_op(Op::HALT, 0);

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
    chunk.emit_op(Op::HALT, 0);

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
    chunk.emit_op(Op::HALT, 0);

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
    chunk.emit_op(Op::HALT, 0);

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
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 3.0);
}

#[test]
fn global_init_runtime_opcode() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    // Pre-set a global via init
    chunk.add_global_init("X", ConstExpr::Value(Value::I32(5)));

    // Also use global_init opcode at runtime to reinit global 0
    chunk.global_inits.push(vybe_bytecode::chunk::GlobalInit {
        name: "Y".into(),
        init: ConstExpr::Add(
            Box::new(ConstExpr::GlobalGet("X".into())),
            Box::new(ConstExpr::Value(Value::I32(10))),
        ),
    });
    chunk.emit_op_u16(Op::GLOBAL_INIT, 1, 0); // init Y (index 1)

    let name = chunk.add_constant(Value::String(Arc::from("Y")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 15);
}

// ── Typed Continuations ─────────────────────────────────────

#[test]
fn cont_new_typed_stores_tag() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    let tag_idx = chunk.add_continuation_tag("generator", "i32", "i32");

    // Create a dummy function (ref_func points to chunk 0)
    chunk.emit_op_u16(Op::REF_FUNC, 0, 0);
    chunk.emit(0, 0); // 0 upvalues

    // cont_new_typed with our tag
    chunk.emit_op_u16(Op::CONT_NEW_TYPED, tag_idx, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    match &result {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            assert_eq!(
                o.properties.get("__cont_tag").unwrap().as_i32(),
                tag_idx as i32
            );
            let yt = o
                .properties
                .get("__cont_yield_type")
                .map(|v| format!("{}", v))
                .unwrap_or_default();
            assert_eq!(yt, "i32");
            let rt = o
                .properties
                .get("__cont_resume_type")
                .map(|v| format!("{}", v))
                .unwrap_or_default();
            assert_eq!(rt, "i32");
        }
        other => panic!("expected Object, got {:?}", other),
    }
}

#[test]
fn suspend_typed_rejects_wrong_yield_type() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    let tag_idx = chunk.add_continuation_tag("yield_i32", "i32", "any");
    let value = chunk.add_constant(Value::String(Arc::from("not an i32")));

    chunk.emit_op_u16(Op::CONST, value, 0);
    chunk.emit_op_u16(Op::SUSPEND_TYPED, tag_idx, 0);
    chunk.emit_op(Op::HALT, 0);

    let err = vm.run(vec![chunk]).expect_err("typed suspend should trap");
    assert!(
        err.to_string()
            .contains("suspend_typed: yield type mismatch"),
        "unexpected error: {err}"
    );
}

#[test]
fn resume_typed_rejects_wrong_resume_type() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;
    let tag_idx = chunk.add_continuation_tag("resume_i32", "any", "i32");
    let resume_value = chunk.add_constant(Value::String(Arc::from("not an i32")));

    chunk.emit_op_u16(Op::REF_FUNC, 0, 0);
    chunk.emit(0, 0);
    chunk.emit_op_u16(Op::CONT_NEW_TYPED, tag_idx, 0);
    chunk.emit_op_u16(Op::CONST, resume_value, 0);
    chunk.emit_op_u16(Op::RESUME_TYPED, tag_idx, 0);
    chunk.emit_op(Op::HALT, 0);

    let err = vm.run(vec![chunk]).expect_err("typed resume should trap");
    assert!(
        err.to_string()
            .contains("resume_typed: resume type mismatch"),
        "unexpected error: {err}"
    );
}

#[test]
fn resume_typed_rejects_wrong_tag() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;
    let actual_tag = chunk.add_continuation_tag("actual", "any", "any");
    let wrong_tag = chunk.add_continuation_tag("wrong", "any", "any");
    let resume_value = chunk.add_constant(Value::I32(1));

    chunk.emit_op_u16(Op::REF_FUNC, 0, 0);
    chunk.emit(0, 0);
    chunk.emit_op_u16(Op::CONT_NEW_TYPED, actual_tag, 0);
    chunk.emit_op_u16(Op::CONST, resume_value, 0);
    chunk.emit_op_u16(Op::RESUME_TYPED, wrong_tag, 0);
    chunk.emit_op(Op::HALT, 0);

    let err = vm
        .run(vec![chunk])
        .expect_err("wrong typed resume tag should trap");
    assert!(
        err.to_string().contains("resume_typed: tag mismatch"),
        "unexpected error: {err}"
    );
}

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

#[test]
fn string_as_ref_passthrough() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    let s = chunk.add_constant(Value::String(Arc::from("hello")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::STRING_AS_REF, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    match &result {
        Value::String(s) => assert_eq!(s.as_ref(), "hello"),
        other => panic!("expected String, got {:?}", other),
    }
}

#[test]
fn string_from_ref_passthrough() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    let s = chunk.add_constant(Value::String(Arc::from("world")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::STRING_AS_REF, 0);
    chunk.emit_op(Op::STRING_FROM_REF, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    match &result {
        Value::String(s) => assert_eq!(s.as_ref(), "world"),
        other => panic!("expected String, got {:?}", other),
    }
}

#[test]
fn string_ref_eq_same_rc() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Push the same constant twice — same Rc
    let s = chunk.add_constant(Value::String(Arc::from("shared")));
    chunk.emit_op_u16(Op::CONST, s, 0);
    chunk.emit_op(Op::DUP, 0);
    chunk.emit_op(Op::STRING_REF_EQ, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Bool(true)));
}

#[test]
fn string_ref_eq_different_rc() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Two different Arc<str> with same content
    let s1 = chunk.add_constant(Value::String(Arc::from("test")));
    let s2 = chunk.add_constant(Value::String(Arc::from("test")));
    chunk.emit_op_u16(Op::CONST, s1, 0);
    chunk.emit_op_u16(Op::CONST, s2, 0);
    chunk.emit_op(Op::STRING_REF_EQ, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    // Different Rc pointers even though same content → false
    assert!(matches!(result, Value::Bool(false)));
}

#[test]
fn string_ref_eq_non_strings() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    let a = chunk.add_constant(Value::I32(42));
    let b = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::STRING_REF_EQ, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Bool(false)));
}
