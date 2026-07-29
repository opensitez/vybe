//! Tests for the function-references proposal: ref.func (0xD2), call_ref (0x14).
//! Covers global-init ConstExpr::RefFunc, runtime REF_FUNC opcode, and call_ref dispatch.

use std::sync::Arc;
use vybe_runtime::chunk::*;
use vybe_runtime::value::ObjectKind;
use vybe_runtime::*;

// ── GlobalInit ConstExpr::RefFunc ─────────────────────────────────────────

#[test]
fn global_init_reffunc_creates_function_object() {
    let mut vm = VM::new();
    let mut script = Chunk::new("<script>");
    script.local_count = 1;
    script.global_inits.push(GlobalInit {
        name: "__fn".to_string(),
        init: ConstExpr::RefFunc(1),
    });
    let name_c = script.add_constant(Value::String(Arc::from("__fn")));
    script.emit_op_u16(opcode::Op::GLOBAL_GET, name_c, 0);
    script.emit_op(opcode::Op::HALT, 0);

    let mut func_chunk = Chunk::new("f");
    func_chunk.arity = 0;
    func_chunk.emit_op(opcode::Op::RETURN, 0);

    let result = vm.run(vec![script, func_chunk]).unwrap();
    match &result {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            assert!(matches!(&o.kind, ObjectKind::Function(_)));
        }
        other => panic!("expected Function, got {:?}", other),
    }
}

#[test]
fn call_ref_invokes_referenced_function() {
    let mut vm = VM::new();
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.global_inits.push(GlobalInit {
        name: "__identity".to_string(),
        init: ConstExpr::RefFunc(1),
    });
    let name_c = script.add_constant(Value::String(Arc::from("__identity")));
    let arg_c = script.add_constant(Value::I32(7));
    script.emit_op_u16(opcode::Op::GLOBAL_GET, name_c, 0);
    script.emit_op_u16(opcode::Op::CONST, arg_c, 0);
    script.emit_op_u8(opcode::Op::CALL_REF, 1, 0);
    script.emit_op(opcode::Op::HALT, 0);

    let mut identity = Chunk::new("identity");
    identity.arity = 1;
    identity.local_count = 1;
    identity.emit_op_u16(opcode::Op::LOCAL_GET, 0, 0);
    identity.emit_op(opcode::Op::RETURN, 0);

    let result = vm.run(vec![script, identity]).unwrap();
    assert_eq!(result.as_i32(), 7);
}

// ── REF_FUNC opcode (runtime closure creation) ────────────────────────────

#[test]
fn ref_func_opcode_registers_in_func_table() {
    let mut vm = VM::new();
    let mut script = Chunk::new("<script>");
    let tidx_key = script.add_constant(Value::String(Arc::from("__table_idx")));

    // REF_FUNC operand: U16 chunk_idx + U8 upvalue_count
    script.emit_op_u16(opcode::Op::REF_FUNC, 1, 0);
    script.emit(0u8, 0); // 0 upvalues

    // STRUCT_GET __table_idx — should be a non-negative integer
    script.emit_op_u16(opcode::Op::STRUCT_GET, tidx_key, 0);
    script.emit_op(opcode::Op::HALT, 0);

    let mut func_chunk = Chunk::new("g");
    func_chunk.arity = 0;
    func_chunk.emit_op(opcode::Op::RETURN, 0);

    let result = vm.run(vec![script, func_chunk]).unwrap();
    assert!(result.as_f64() >= 0.0, "table_idx should be non-negative");
}

#[test]
fn call_ref_with_multiple_args() {
    let mut vm = VM::new();
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.global_inits.push(GlobalInit {
        name: "__add".to_string(),
        init: ConstExpr::RefFunc(1),
    });
    let name_c = script.add_constant(Value::String(Arc::from("__add")));
    let a_c = script.add_constant(Value::I32(10));
    let b_c = script.add_constant(Value::I32(32));
    script.emit_op_u16(opcode::Op::GLOBAL_GET, name_c, 0);
    script.emit_op_u16(opcode::Op::CONST, a_c, 0);
    script.emit_op_u16(opcode::Op::CONST, b_c, 0);
    script.emit_op_u8(opcode::Op::CALL_REF, 2, 0);
    script.emit_op(opcode::Op::HALT, 0);

    let mut add_fn = Chunk::new("add");
    add_fn.arity = 2;
    add_fn.local_count = 2;
    add_fn.emit_op_u16(opcode::Op::LOCAL_GET, 0, 0);
    add_fn.emit_op_u16(opcode::Op::LOCAL_GET, 1, 0);
    add_fn.emit_op(opcode::Op::I32_ADD, 0);
    add_fn.emit_op(opcode::Op::RETURN, 0);

    let result = vm.run(vec![script, add_fn]).unwrap();
    assert_eq!(result.as_i32(), 42);
}

#[test]
fn call_ref_non_function_traps() {
    let mut vm = VM::new();
    let mut script = Chunk::new("<script>");
    let k = script.add_constant(Value::I32(123));
    script.emit_op_u16(opcode::Op::CONST, k, 0);
    script.emit_op_u8(opcode::Op::CALL_REF, 0, 0);
    script.emit_op(opcode::Op::HALT, 0);

    let err = vm.run(vec![script]).unwrap_err().to_string();
    assert!(err.contains("call") || err.contains("function"));
}
