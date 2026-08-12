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

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn add_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str) -> u16 {
    chunks[current].add_import(module, name)
}

/// `new Proxy(target, handler)`. Stack: [target, handler] → [proxy_obj].
pub fn emit_proxy_create(chunks: &mut [Chunk], current: usize, line: u32) {
    let new_idx = add_import(chunks, current, "ecma:proxy", "new");
    let chunk = &mut chunks[current];
    chunk.emit_call(new_idx, 2, line);
}

/// Proxy `get` trap dispatch. Stack: [obj, key] → [value].
pub fn emit_proxy_get_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let get_idx = add_import(chunks, current, "ecma:proxy", "get");
    let chunk = &mut chunks[current];
    chunk.emit_call(get_idx, 2, line);
}

/// Proxy `set` trap dispatch. Stack: [obj, key, value] → [value].
pub fn emit_proxy_set_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let set_idx = add_import(chunks, current, "ecma:proxy", "set");
    let chunk = &mut chunks[current];
    let value_local = alloc_local(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, value_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_local, line);
    chunk.emit_call(set_idx, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_local, line);
}

/// Strict-mode Proxy `set` dispatch. Stack: [obj, key, value] → [bool] —
/// the [[Set]] success flag is LEFT on the stack so the caller can apply
/// the §13.15.2 strict-assignment TypeError check.
pub fn emit_proxy_set_dispatch_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    let set_idx = add_import(chunks, current, "ecma:proxy", "set");
    let chunk = &mut chunks[current];
    chunk.emit_call(set_idx, 3, line);
}

/// Proxy `has` trap dispatch. Stack: [obj, key] → [bool]. Used by `in`.
pub fn emit_proxy_has_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let has_idx = add_import(chunks, current, "ecma:proxy", "has");
    let chunk = &mut chunks[current];
    chunk.emit_call(has_idx, 2, line);
}

/// Proxy `deleteProperty` trap dispatch. Stack: [obj, key] → [bool].
pub fn emit_proxy_delete_property_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let delete_idx = add_import(chunks, current, "ecma:proxy", "deleteProperty");
    let chunk = &mut chunks[current];
    chunk.emit_call(delete_idx, 2, line);
}

/// Proxy `apply` trap dispatch. Stack: [callable_or_proxy, this_arg, args_array] → [result].
pub fn emit_proxy_apply_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let apply_idx = add_import(chunks, current, "ecma:proxy", "apply");
    let chunk = &mut chunks[current];
    chunk.emit_call(apply_idx, 3, line);
}

/// Proxy `ownKeys` trap dispatch. Stack: [obj] → [keys_array].
pub fn emit_proxy_own_keys_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = add_import(chunks, current, "ecma:proxy", "ownKeys");
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, 1, line);
}

/// Proxy `getOwnPropertyDescriptor` trap dispatch. Stack: [obj, key] → [descriptor|undefined].
pub fn emit_proxy_get_own_property_descriptor_dispatch(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
) {
    let idx = add_import(chunks, current, "ecma:proxy", "getOwnPropertyDescriptor");
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, 2, line);
}

/// Proxy `defineProperty` trap dispatch. Stack: [obj, key, descriptor] → [bool].
pub fn emit_proxy_define_property_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = add_import(chunks, current, "ecma:proxy", "defineProperty");
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, 3, line);
}

/// Proxy `getPrototypeOf` trap dispatch. Stack: [obj] → [prototype|null].
pub fn emit_proxy_get_prototype_of_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = add_import(chunks, current, "ecma:proxy", "getPrototypeOf");
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, 1, line);
}

/// Proxy `setPrototypeOf` trap dispatch. Stack: [obj, prototype|null] → [bool].
pub fn emit_proxy_set_prototype_of_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = add_import(chunks, current, "ecma:proxy", "setPrototypeOf");
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, 2, line);
}

/// Proxy `isExtensible` trap dispatch. Stack: [obj] → [bool].
pub fn emit_proxy_is_extensible_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = add_import(chunks, current, "ecma:proxy", "isExtensible");
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, 1, line);
}

/// Proxy `preventExtensions` trap dispatch. Stack: [obj] → [bool].
pub fn emit_proxy_prevent_extensions_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = add_import(chunks, current, "ecma:proxy", "preventExtensions");
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, 1, line);
}

/// Proxy `construct` trap dispatch. Stack: [constructor_or_proxy, args_array] → [object].
pub fn emit_proxy_construct_dispatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = add_import(chunks, current, "ecma:proxy", "construct");
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, 2, line);
}
