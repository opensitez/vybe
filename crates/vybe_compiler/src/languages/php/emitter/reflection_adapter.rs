//! PHP Reflection classes — Rust inline opcode emitters.
//!
//! PHP-over-JS: `Reflection*` are thin objects whose methods compose the
//! SAME introspection JS uses — `ecma:reflect.get`/`set` for property
//! access, dynamic method refs for `invoke`. No PHP-only metadata store,
//! no host fns, no common/VM changes. Same shape as `spl_adapter.rs`.
//!
//! The walker rewrites `new ReflectionClass($n)` → `__refl_class($n)` etc.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn sconst(c: &mut Chunk, s: &str) -> u16 {
    c.add_constant(Value::String(Arc::from(s)))
}

/// `() -> this.<field>`.
fn build_field_getter(chunks: &mut Vec<Chunk>, name: &str, field: &str, line: u32) -> usize {
    let mut c = Chunk::new(name);
    c.arity = 1;
    let k = sconst(&mut c, field);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, k, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `getValue($obj)` → `Reflect.get($obj, this.<field>)`.
fn build_reflect_get(chunks: &mut Vec<Chunk>, name: &str, field: &str, line: u32) -> usize {
    let mut c = Chunk::new(name);
    c.arity = 2;
    let k = sconst(&mut c, field);
    let get_i = c.add_import("ecma:reflect".to_string(), "get".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 1, line); // obj
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, k, line); // this.field
    c.emit_op_u16(Op::CALL_IMPORT, get_i, line);
    c.emit(2, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// `setValue($obj, $v)` → `Reflect.set($obj, this.<field>, $v)`.
fn build_reflect_set(chunks: &mut Vec<Chunk>, name: &str, field: &str, line: u32) -> usize {
    let mut c = Chunk::new(name);
    c.arity = 3;
    let k = sconst(&mut c, field);
    let set_i = c.add_import("ecma:reflect".to_string(), "set".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 1, line); // obj
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, k, line); // this.field
    c.emit_op_u16(Op::LOCAL_GET, 2, line); // value
    c.emit_op_u16(Op::CALL_IMPORT, set_i, line);
    c.emit(3, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(3);
    chunks.push(c);
    chunks.len() - 1
}

/// `invoke($obj, $arg)` → `Reflect.get($obj, this.method)($obj, $arg)`.
fn build_method_invoke(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__refl_invoke");
    c.arity = 3; // this, obj, arg1
    let method_k = sconst(&mut c, "method");
    let get_i = c.add_import("ecma:reflect".to_string(), "get".to_string());
    // method ref = Reflect.get(obj, this.method)
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, method_k, line);
    c.emit_op_u16(Op::CALL_IMPORT, get_i, line);
    c.emit(2, line);
    // call method(obj, arg1)
    c.emit_op_u16(Op::LOCAL_GET, 1, line); // obj = this for the method
    c.emit_op_u16(Op::LOCAL_GET, 2, line); // arg1
    c.emit_op_u8(Op::CALL_REF, 2, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(3);
    chunks.push(c);
    chunks.len() - 1
}

/// Stamp `__type`, set each `(field, slot)`, bind each `(method, idx)`,
/// leave the instance on the stack.
fn finish(
    chunk: &mut Chunk,
    this_slot: u16,
    kind: &str,
    fields: &[(&str, u16)],
    binds: &[(&str, usize)],
    line: u32,
) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op(Op::DROP, line);
    // __type
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let kc = sconst(chunk, kind);
    chunk.emit_op_u16(Op::CONST, kc, line);
    let tk = sconst(chunk, "__type");
    chunk.emit_op_u16(Op::STRUCT_SET, tk, line);
    chunk.emit_op(Op::DROP, line);
    // fields
    for (fname, fslot) in fields {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, *fslot, line);
        let fk = sconst(chunk, fname);
        chunk.emit_op_u16(Op::STRUCT_SET, fk, line);
        chunk.emit_op(Op::DROP, line);
    }
    // methods
    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, *midx as u16, line);
        chunk.emit(0, line);
        let mk = sconst(chunk, mname);
        chunk.emit_op_u16(Op::STRUCT_SET, mk, line);
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

/// `new ReflectionClass($name)`.
pub fn emit_refl_class(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let getname = build_field_getter(chunks, "__refl_class_name", "name", line);
    let chunk = &mut chunks[current];
    let name_slot = chunk.local_count;
    let this_slot = chunk.local_count + 1;
    chunk.local_count += 2;
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_op(Op::DROP, line);
    finish(
        chunk,
        this_slot,
        "ReflectionClass",
        &[("name", name_slot)],
        &[("getname", getname)],
        line,
    );
}

/// `new ReflectionMethod($class, $method)`.
pub fn emit_refl_method(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let getname = build_field_getter(chunks, "__refl_method_name", "method", line);
    let invoke = build_method_invoke(chunks, line);
    let chunk = &mut chunks[current];
    let method_slot = chunk.local_count;
    let class_slot = chunk.local_count + 1;
    let this_slot = chunk.local_count + 2;
    chunk.local_count += 3;
    // args: [class, method], method on top
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, class_slot, line);
    chunk.emit_op(Op::DROP, line);
    finish(
        chunk,
        this_slot,
        "ReflectionMethod",
        &[("method", method_slot), ("class", class_slot)],
        &[("getname", getname), ("invoke", invoke)],
        line,
    );
}

/// `new ReflectionProperty($class, $prop)`.
pub fn emit_refl_property(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let getname = build_field_getter(chunks, "__refl_prop_name", "prop", line);
    let getvalue = build_reflect_get(chunks, "__refl_getvalue", "prop", line);
    let setvalue = build_reflect_set(chunks, "__refl_setvalue", "prop", line);
    let chunk = &mut chunks[current];
    let prop_slot = chunk.local_count;
    let class_slot = chunk.local_count + 1;
    let this_slot = chunk.local_count + 2;
    chunk.local_count += 3;
    chunk.emit_op_u16(Op::LOCAL_SET, prop_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, class_slot, line);
    chunk.emit_op(Op::DROP, line);
    finish(
        chunk,
        this_slot,
        "ReflectionProperty",
        &[("prop", prop_slot), ("class", class_slot)],
        &[
            ("getname", getname),
            ("getvalue", getvalue),
            ("setvalue", setvalue),
        ],
        line,
    );
}
