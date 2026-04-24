//! RefFunc + call_ref dispatch.
//!
//! Tests that a global initialised to `RefFunc(idx)` resolves to a
//! Function value and that `call_ref` dispatches to the referenced
//! chunk. No dynamic-array work here — just the RefFunc/CallRef
//! plumbing. The earlier version of this test built a `range()`
//! stdlib chunk that pushed values via the removed `ARRAY_PUSH`
//! opcode; that behavioural surface now lives under
//! `vybe_host::vybe:js-array.*` (see
//! `crates/vybe_host/tests/js_builtins_behavior_test.rs`). Here we
//! keep the test focused on the single concern it names.

use vybe_bytecode::*;
use vybe_bytecode::chunk::*;
use std::sync::Arc;

#[test]
fn global_init_reffunc_then_callref() {
    let mut vm = VM::new();
    let mut script = Chunk::new("<script>");
    script.local_count = 2;

    // Global init: __test = RefFunc(1) → Function ref to chunk 1.
    script.global_inits.push(GlobalInit {
        name: "__test".to_string(),
        init: ConstExpr::RefFunc(1),
    });

    // Call __test(7) and return its result.
    let name_c = script.add_constant(Value::String(Arc::from("__test")));
    let c7 = script.add_constant(Value::I32(7));
    script.emit_op_u16(opcode::Op::GLOBAL_GET, name_c, 0);
    script.emit_op_u16(opcode::Op::CONST, c7, 0);
    script.emit_op_u8(opcode::Op::CALL_REF, 1, 0);
    script.emit_op(opcode::Op::HALT, 0);

    // Chunk 1: identity — returns its arg unchanged.
    let mut identity = Chunk::new("identity");
    identity.arity = 1;
    identity.local_count = 1; // slot 0 = arg (WASM convention)
    identity.emit_op_u16(opcode::Op::LOCAL_GET, 0, 0);
    identity.emit_op(opcode::Op::RETURN, 0);

    let result = vm.run(vec![script, identity]).unwrap();
    assert_eq!(result.as_i32(), 7, "RefFunc global → call_ref should return identity(7) = 7");
}
