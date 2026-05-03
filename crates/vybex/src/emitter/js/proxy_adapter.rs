//! ECMA-262 §10.5 Proxy — inline opcode emitters.
//!
//! Proxy is implemented as compile-time inline dispatch:
//!
//! 1. **`new Proxy(target, handler)`** — `emit_proxy_create`. Result is
//!    an Ordinary object stamped with `__vybe_proxy_target` and
//!    `__vybe_proxy_handler` properties. Per ECMA-262 §10.5.2 the
//!    proxy is an exotic object whose [[Prototype]] reflects the
//!    target — we keep it Ordinary so existing `STRUCT_GET` /
//!    `STRUCT_SET` paths see real properties.
//!
//! 2. **Member / Index read** — `emit_proxy_get_dispatch`. Inline
//!    structured-control-flow dispatch:
//!
//!      handler = obj.__vybe_proxy_handler
//!      if handler is null: → fall back to ARRAY_GET
//!      trap = handler.get
//!      if trap is not function: → ARRAY_GET on the underlying target
//!      bind __js_this = handler; result = handler.get(target, key, obj)
//!
//! 3. **Member / Index assign** — `emit_proxy_set_dispatch`. Same shape
//!    with the `set` trap and (target, key, value, obj) args.
//!
//! 4. **`prop in obj`** — `emit_proxy_has_dispatch`. Calls `handler.has`
//!    if defined; else falls through to `ecma:object.hasIn`.
//!
//! Modules without `Proxy` keep direct `STRUCT_GET` / `ARRAY_GET` —
//! zero overhead. The compiler scans the AST once and sets
//! `Compiler::uses_proxy` for the routing decision.

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

const HANDLER_KEY: &str = "__vybe_proxy_handler";
const TARGET_KEY: &str = "__vybe_proxy_target";

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}

fn add_import(chunks: &mut [Chunk], module: &str, name: &str) -> u16 {
    chunks[0].add_import(module, name)
}

/// `new Proxy(target, handler)` — stamp a wrapper Ordinary object with
/// the two well-known proxy properties. Stack: [target, handler] →
/// [proxy_obj]. We use `STRUCT_NEW 0` (Ordinary) so existing dispatch
/// paths see real properties rather than gating on a type_id.
pub fn emit_proxy_create(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let target_local = alloc_local(chunk);
    let handler_local = alloc_local(chunk);
    let wrapper_local = alloc_local(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, handler_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_local, line);  chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, wrapper_local, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, wrapper_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_local, line);
    let target_key = chunk.add_constant(Value::String(Arc::from(TARGET_KEY)));
    chunk.emit_op_u16(Op::STRUCT_SET, target_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, wrapper_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    let handler_key = chunk.add_constant(Value::String(Arc::from(HANDLER_KEY)));
    chunk.emit_op_u16(Op::STRUCT_SET, handler_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, wrapper_local, line);
}

/// Proxy `get` trap dispatch. Stack: [obj, key] → [value].
pub fn emit_proxy_get_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_local = alloc_local(chunk);
    let key_local = alloc_local(chunk);
    let handler_local = alloc_local(chunk);
    let trap_local = alloc_local(chunk);
    let target_local = alloc_local(chunk);
    let saved_this_local = alloc_local(chunk);
    let result_local = alloc_local(chunk);

    let handler_key = chunk.add_constant(Value::String(Arc::from(HANDLER_KEY)));
    let target_key = chunk.add_constant(Value::String(Arc::from(TARGET_KEY)));
    let get_key = chunk.add_constant(Value::String(Arc::from("get")));
    let func_str = chunk.add_constant(Value::String(Arc::from("function")));
    let js_this = chunk.add_constant(Value::String(Arc::from("__js_this")));

    chunk.emit_op_u16(Op::LOCAL_SET, key_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_local, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::UNDEFINED, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line); chunk.emit_op(Op::DROP, line);

    let exit_block = chunk.emit_block(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, handler_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, handler_local, line); chunk.emit_op(Op::DROP, line);

    let no_handler = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::DYN_NOT, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line); chunk.patch_block(no_handler);

    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, get_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, trap_local, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, target_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_local, line); chunk.emit_op(Op::DROP, line);

    let no_trap = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, trap_local, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    chunk.emit_op_u16(Op::CONST, func_str, line);
    chunk.emit_op(Op::DYN_EQ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line); chunk.patch_block(no_trap);

    chunk.emit_op_u16(Op::GLOBAL_GET, js_this, line);
    chunk.emit_op_u16(Op::LOCAL_SET, saved_this_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, js_this, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, trap_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u8(Op::CALL_REF, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, saved_this_local, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, js_this, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line); chunk.patch_block(exit_block);

    chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
}

/// Proxy `set` trap dispatch. Stack: [obj, key, value] → [value].
pub fn emit_proxy_set_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_local = alloc_local(chunk);
    let key_local = alloc_local(chunk);
    let value_local = alloc_local(chunk);
    let handler_local = alloc_local(chunk);
    let trap_local = alloc_local(chunk);
    let target_local = alloc_local(chunk);
    let saved_this_local = alloc_local(chunk);

    let handler_key = chunk.add_constant(Value::String(Arc::from(HANDLER_KEY)));
    let target_key = chunk.add_constant(Value::String(Arc::from(TARGET_KEY)));
    let set_key = chunk.add_constant(Value::String(Arc::from("set")));
    let func_str = chunk.add_constant(Value::String(Arc::from("function")));
    let js_this = chunk.add_constant(Value::String(Arc::from("__js_this")));

    chunk.emit_op_u16(Op::LOCAL_SET, value_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, key_local, line);   chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_local, line);   chunk.emit_op(Op::DROP, line);

    let exit_block = chunk.emit_block(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, handler_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, handler_local, line); chunk.emit_op(Op::DROP, line);

    let no_handler = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::DYN_NOT, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_local, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line); chunk.patch_block(no_handler);

    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, set_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, trap_local, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, target_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_local, line); chunk.emit_op(Op::DROP, line);

    let no_trap = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, trap_local, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    chunk.emit_op_u16(Op::CONST, func_str, line);
    chunk.emit_op(Op::DYN_EQ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_local, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line); chunk.patch_block(no_trap);

    chunk.emit_op_u16(Op::GLOBAL_GET, js_this, line);
    chunk.emit_op_u16(Op::LOCAL_SET, saved_this_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, js_this, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, trap_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u8(Op::CALL_REF, 4, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, saved_this_local, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, js_this, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line); chunk.patch_block(exit_block);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, value_local, line);
}

/// Proxy `has` trap dispatch. Stack: [obj, key] → [bool]. Used by `in`.
pub fn emit_proxy_has_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let has_in_idx = add_import(chunks, "ecma:object", "hasIn");
    let chunk = &mut chunks[current];
    let obj_local = alloc_local(chunk);
    let key_local = alloc_local(chunk);
    let handler_local = alloc_local(chunk);
    let trap_local = alloc_local(chunk);
    let target_local = alloc_local(chunk);
    let saved_this_local = alloc_local(chunk);
    let result_local = alloc_local(chunk);

    let handler_key = chunk.add_constant(Value::String(Arc::from(HANDLER_KEY)));
    let target_key = chunk.add_constant(Value::String(Arc::from(TARGET_KEY)));
    let has_key = chunk.add_constant(Value::String(Arc::from("has")));
    let func_str = chunk.add_constant(Value::String(Arc::from("function")));
    let js_this = chunk.add_constant(Value::String(Arc::from("__js_this")));

    chunk.emit_op_u16(Op::LOCAL_SET, key_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_local, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::FALSE, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line); chunk.emit_op(Op::DROP, line);

    let exit_block = chunk.emit_block(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, handler_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, handler_local, line); chunk.emit_op(Op::DROP, line);

    let no_handler = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::DYN_NOT, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, has_in_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line); chunk.patch_block(no_handler);

    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, has_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, trap_local, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, target_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_local, line); chunk.emit_op(Op::DROP, line);

    let no_trap = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, trap_local, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    chunk.emit_op_u16(Op::CONST, func_str, line);
    chunk.emit_op(Op::DYN_EQ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, has_in_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line); chunk.patch_block(no_trap);

    chunk.emit_op_u16(Op::GLOBAL_GET, js_this, line);
    chunk.emit_op_u16(Op::LOCAL_SET, saved_this_local, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, handler_local, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, js_this, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, trap_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_local, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, saved_this_local, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, js_this, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line); chunk.patch_block(exit_block);

    chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
}
