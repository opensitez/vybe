use crate::emitter::instructions::core_wasm;
use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::classes::{
    emit_bind_bound_method_with_aliases, emit_bind_getter, emit_bind_setter,
};
use crate::emitter::functions::create_function_chunk;

const SERIAL_KIND_KEY: &str = "vybe$php_ser_kind";

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_op(Op::NULL, line),
        Value::BigInt(v) => chunk.emit_i64_const(*v, line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),
        
        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

fn call_import_into(
    _imports: &mut Chunk,
    code: &mut Chunk,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = code.add_import(module.to_string(), name.to_string());
    code.emit_call(idx, argc, line);
}

fn call_ref(chunk: &mut Chunk, argc: u8, line: u32) {
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(argc, line);
}

fn ref_func(chunk: &mut Chunk, func_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_idx as u16, line);
    chunk.emit(0, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn dynamic_get_from_slots(chunk: &mut Chunk, obj_slot: u16, key_slot: u16, line: u32) {
    lget(chunk, obj_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn dynamic_set_from_slots(
    chunk: &mut Chunk,
    obj_slot: u16,
    key_slot: u16,
    value_slot: u16,
    line: u32,
) {
    lget(chunk, obj_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
}

fn set_struct_from_slot(chunk: &mut Chunk, obj_slot: u16, key: &str, value_slot: u16, line: u32) {
    lget(chunk, obj_slot, line);
    lget(chunk, value_slot, line);
    struct_set_key(chunk, key, line);
}

pub fn emit_php_header(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let send_header_import =
        chunks[0].add_import("node:http".to_string(), "send_header_raw".to_string());
    let response_code_import =
        chunks[0].add_import("node:http".to_string(), "http_response_code".to_string());
    let string_import = chunks[0].add_import("ecma:string".to_string(), "String".to_string());
    let lower_import = chunks[0].add_import("ecma:string".to_string(), "toLowerCase".to_string());
    let starts_with_import =
        chunks[0].add_import("ecma:string".to_string(), "startsWith".to_string());
    let chunk = &mut chunks[current];

    let header_slot = alloc_local(chunk);
    let replace_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let response_code_slot = if argc >= 3 {
        Some(alloc_local(chunk))
    } else {
        None
    };

    if let Some(slot) = response_code_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = replace_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, header_slot, line);

    lget(chunk, header_slot, line);
    if let Some(slot) = replace_slot {
        lget(chunk, slot, line);
    }
    if let Some(slot) = response_code_slot {
        lget(chunk, slot, line);
    }
    chunk.emit_call(send_header_import, argc, line);
    chunk.emit_op(Op::DROP, line);

    if let Some(slot) = response_code_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(0.0), line);
        crate::emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
    }

    lget(chunk, header_slot, line);
    chunk.emit_call(string_import, 1, line);
    chunk.emit_call(lower_import, 1, line);
    push_str(chunk, "location:", line);
    chunk.emit_call(starts_with_import, 2, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_call(response_code_import, 0, line);
    push_const(chunk, Value::F64(200.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);

    push_const(chunk, Value::F64(302.0), line);
    chunk.emit_call(response_code_import, 1, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    if response_code_slot.is_some() {
        chunk.emit_end(line);
    }

    chunk.emit_op(Op::NULL, line);
}

pub fn emit_php_extension_loaded(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let string_import = chunks[0].add_import("ecma:string".to_string(), "String".to_string());
    let lower_import = chunks[0].add_import("ecma:string".to_string(), "toLowerCase".to_string());
    let chunk = &mut chunks[current];
    if argc == 0 {
        push_const(chunk, Value::Bool(false), line);
        return;
    }

    // extension_loaded() is unary; discard extra args defensively.
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_call(string_import, 1, line);
    chunk.emit_call(lower_import, 1, line);

    let ext_slot = alloc_local(chunk);
    lset(chunk, ext_slot, line);

    lget(chunk, ext_slot, line);
    push_str(chunk, "mysqli", line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, ext_slot, line);
    push_str(chunk, "mysqlnd", line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, ext_slot, line);
    push_str(chunk, "pdo_mysql", line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, ext_slot, line);
    push_str(chunk, "mysql", line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_php_phpversion(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "8.0.0", line);
}

pub fn emit_php_spl_autoload_register(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_spl_autoload_unregister(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_session_start(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let set_cookie_import = chunks[0].add_import("node:http".to_string(), "set_cookie".to_string());
    let chunk = &mut chunks[current];
    let needs_cookie = chunk.add_constant(Value::String(Arc::from("__php_session_needs_cookie")));
    let session_id = chunk.add_constant(Value::String(Arc::from("__php_session_id")));
    let started = chunk.add_constant(Value::String(Arc::from("__php_session_started")));
    let destroyed = chunk.add_constant(Value::String(Arc::from("__php_session_destroyed")));

    push_const(chunk, Value::Bool(true), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, started, line);
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, destroyed, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::GLOBAL_GET, needs_cookie, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    push_str(chunk, "PHPSESSID", line);
    chunk.emit_op_u16(Op::GLOBAL_GET, session_id, line);
    chunk.emit_call(set_cookie_import, 2, line);
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, needs_cookie, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_session_unset(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    let session = chunk.add_constant(Value::String(Arc::from("$_SESSION")));
    chunk.emit_op_u16(Op::GLOBAL_SET, session, line);
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_session_destroy(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_php_session_unset(chunks, current, 0, line);
    let set_cookie_import = chunks[0].add_import("node:http".to_string(), "set_cookie".to_string());
    let chunk = &mut chunks[current];
    let started = chunk.add_constant(Value::String(Arc::from("__php_session_started")));
    let destroyed = chunk.add_constant(Value::String(Arc::from("__php_session_destroyed")));
    let needs_cookie = chunk.add_constant(Value::String(Arc::from("__php_session_needs_cookie")));

    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, started, line);
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, destroyed, line);
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, needs_cookie, line);
    chunk.emit_op(Op::DROP, line);

    push_str(chunk, "PHPSESSID", line);
    push_str(chunk, "", line);
    chunk.emit_call(set_cookie_import, 2, line);
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(true), line);
}

fn helper_loop_start(chunk: &mut Chunk, line: u32) -> crate::emitter::loops::LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    crate::emitter::loops::LoopState {
        block_patch,
        loop_patch,
        body_block_patch: None,
    }
}

fn helper_loop_cond(chunk: &mut Chunk, line: u32) {
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
}

fn helper_loop_end(chunk: &mut Chunk, state: crate::emitter::loops::LoopState, line: u32) {
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(state.block_patch);
}

fn emit_nullish_return(chunk: &mut Chunk, value_slot: u16, line: u32) {
    lget(chunk, value_slot, line);
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_else(line);
    { let undef_idx = chunk.add_import("wasm:js-undefined", "test"); chunk.emit_call(undef_idx, 1, line); }
    chunk.emit_if(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_is_array_into(imports: &mut Chunk, code: &mut Chunk, value_slot: u16, line: u32) {
    lget(code, value_slot, line);
    call_import_into(imports, code, "ecma:array", "isArray", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(code, line);
}

fn bump_loop_index(chunk: &mut Chunk, i_slot: u16, line: u32) {
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
}

fn build_php_alloc_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let types = chunks[0].types.clone();

    let mut helper = create_function_chunk("__php_unserialize_alloc", 1);
    helper.local_count = 1;
    let class_slot = 0;
    let obj_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        emit_nullish_return(&mut helper, class_slot, line);

        for ty in types.iter().filter(|ty| !ty.is_interface) {
            lget(&mut helper, class_slot, line);
            push_str(&mut helper, &ty.name, line);
            crate::emitter::ops::emit_dyn_eq(&mut helper, line);
            helper.emit_if(line);

            helper.emit_op_u16(Op::STRUCT_NEW, 0, line);
            lset(&mut helper, obj_slot, line);

            lget(&mut helper, obj_slot, line);
            push_str(&mut helper, &ty.name, line);
            struct_set_key(&mut helper, "__type", line);

            lget(&mut helper, obj_slot, line);
            push_str(&mut helper, &ty.name.to_lowercase(), line);
            struct_set_key(&mut helper, "__control_name", line);

            let tid_name =
                helper.add_constant(Value::String(Arc::from(format!("__tid_{}", ty.name))));
            lget(&mut helper, obj_slot, line);
            helper.emit_op_u16(Op::GLOBAL_GET, tid_name, line);
            let tid_key = helper.add_constant(Value::String(Arc::from("__type_id")));
            helper.emit_op_u16(Op::STRUCT_SET, tid_key, line);
            helper.emit_op(Op::DROP, line);

            for field in &ty.fields {
                lget(&mut helper, obj_slot, line);
                helper.emit_op(Op::NULL, line);
                struct_set_key(&mut helper, field, line);
            }

            for (method_name, method_chunk_idx) in &ty.methods {
                if method_name.starts_with("__get_") {
                    let prop = method_name
                        .strip_prefix("__get_")
                        .unwrap_or(method_name.as_str());
                    emit_bind_getter(&mut helper, obj_slot, prop, *method_chunk_idx, line);
                } else if method_name.starts_with("__set_") {
                    let prop = method_name
                        .strip_prefix("__set_")
                        .unwrap_or(method_name.as_str());
                    emit_bind_setter(&mut helper, obj_slot, prop, *method_chunk_idx, line);
                } else {
                    emit_bind_bound_method_with_aliases(
                        &mut helper,
                        obj_slot,
                        method_name,
                        *method_chunk_idx,
                        None,
                        line,
                    );
                }
            }

            lget(&mut helper, obj_slot, line);
            helper.emit_op(Op::RETURN, line);
            helper.emit_end(line);
        }

        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        helper.emit_op(Op::RETURN, line);
    }

    chunks.push(helper);
    helper_idx
}

/// Recursive PHP value → JSON-serializable shape. An associative array is an
/// `ObjectKind::Map`, and `ecma:json.stringify` renders a bare Map as `{}` (ECMA
/// §25 — `JSON.stringify(new Map())` is `{}`), so PHP must convert Map → plain
/// Object first. Arrays recurse on elements; nested Maps recurse too. Key order
/// is the Map's native (`ecma:object.keys`) insertion order — no `__keys`/CSV
/// side-band. The helper self-recurses via its own func ref.
fn build_php_json_normalize_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut helper = create_function_chunk("__php_json_normalize", 1);
    helper.local_count = 1;

    let value_slot = 0u16;
    let out_slot = alloc_local(&mut helper);
    let keys_slot = alloc_local(&mut helper);
    let key_slot = alloc_local(&mut helper);
    let i_slot = alloc_local(&mut helper);
    let n_slot = alloc_local(&mut helper);
    let _type_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        // ALL imports in a helper chunk must register on `imports` (chunks[0])
        // via the `_into` ops — `add_import` is per-chunk and emit_call
        // resolves against chunks[0]'s table, so the non-`_into` ops (which use
        // the helper's own list) produce clashing indices.

        // null / undefined → pass through (stringify handles them).
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        lget(&mut helper, value_slot, line);
        { let undef_idx = helper.add_import("wasm:js-undefined", "test"); helper.emit_call(undef_idx, 1, line); }
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        // Sequential array → new array of normalized elements.
        lget(&mut helper, value_slot, line);
        call_import_into(imports, &mut helper, "ecma:array", "isArray", 1, line);
        crate::emitter::ops::emit_dyn_to_bool_into(imports, &mut helper, line);
        helper.emit_if(line);
        helper.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let arr_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt_into(imports, &mut helper, line);
        crate::emitter::ops::emit_dyn_to_bool_into(imports, &mut helper, line);
        crate::emitter::ops::emit_dyn_not_into(imports, &mut helper, line);
        helper.emit_br_if(1, line);
        lget(&mut helper, out_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, value_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        call_ref(&mut helper, 1, line);
        call_import_into(imports, &mut helper, "ecma:array", "push", 2, line);
        helper.emit_op(Op::DROP, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, arr_loop, line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        // Object / Map → fromEntries of [k, normalize(v[k])] in native key order.
        // object test: not null AND not number AND not string AND not boolean
        let obj_norm_slot = alloc_local(&mut helper);
        lget(&mut helper, value_slot, line);
        helper.emit_op_u16(Op::LOCAL_SET, obj_norm_slot, line);
        helper.emit_op(Op::DROP, line);
        // not null
        lget(&mut helper, obj_norm_slot, line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_op(Op::I32_EQZ, line);
        // AND not number
        lget(&mut helper, obj_norm_slot, line);
        let test_num_norm = helper.add_import("wasm:js-number", "test");
        helper.emit_call(test_num_norm, 1, line);
        helper.emit_op(Op::I32_EQZ, line);
        helper.emit_op(Op::I32_AND, line);
        // AND not string
        lget(&mut helper, obj_norm_slot, line);
        let test_str_norm = helper.add_import("wasm:js-string", "test");
        helper.emit_call(test_str_norm, 1, line);
        helper.emit_op(Op::I32_EQZ, line);
        helper.emit_op(Op::I32_AND, line);
        // AND not boolean
        lget(&mut helper, obj_norm_slot, line);
        let test_bool_norm = helper.add_import("wasm:js-boolean", "test");
        helper.emit_call(test_bool_norm, 1, line);
        helper.emit_op(Op::I32_EQZ, line);
        helper.emit_op(Op::I32_AND, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "keys", 1, line);
        lset(&mut helper, keys_slot, line);
        helper.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, keys_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let obj_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt_into(imports, &mut helper, line);
        crate::emitter::ops::emit_dyn_to_bool_into(imports, &mut helper, line);
        crate::emitter::ops::emit_dyn_not_into(imports, &mut helper, line);
        helper.emit_br_if(1, line);
        lget(&mut helper, keys_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        // pair = [ key, normalize(value[key]) ] ; out.push(pair)
        lget(&mut helper, out_slot, line);
        lget(&mut helper, key_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, value_slot, line);
        lget(&mut helper, key_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        call_ref(&mut helper, 1, line);
        helper.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, line);
        call_import_into(imports, &mut helper, "ecma:array", "push", 2, line);
        helper.emit_op(Op::DROP, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, obj_loop, line);
        lget(&mut helper, out_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "fromEntries", 1, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        // Primitive (boolean / number / string) → pass through.
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
    }

    chunks.push(helper);
    helper_idx
}

/// Build the normalizer and call it on `value_slot`, leaving the
/// JSON-serializable (Map-free) value on the stack.
pub fn emit_php_json_normalize(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    line: u32,
) {
    let helper_idx = build_php_json_normalize_helper(chunks, line);
    let chunk = &mut chunks[current];
    ref_func(chunk, helper_idx, line);
    lget(chunk, value_slot, line);
    call_ref(chunk, 1, line);
}

fn build_php_serialize_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut helper = create_function_chunk("__php_serialize_value", 1);
    helper.local_count = 1;

    let value_slot = 0;
    let _type_slot = alloc_local(&mut helper);
    let out_slot = alloc_local(&mut helper);
    let items_slot = alloc_local(&mut helper);
    let assoc_slot = alloc_local(&mut helper);
    let names_slot = alloc_local(&mut helper);
    let key_slot = alloc_local(&mut helper);
    let tmp_slot = alloc_local(&mut helper);
    let i_slot = alloc_local(&mut helper);
    let n_slot = alloc_local(&mut helper);
    let method_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        emit_nullish_return(&mut helper, value_slot, line);

        // boolean test
        lget(&mut helper, value_slot, line);
        let test_bool_ser = helper.add_import("wasm:js-boolean", "test");
        helper.emit_call(test_bool_ser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        // number test
        lget(&mut helper, value_slot, line);
        let test_num_ser = helper.add_import("wasm:js-number", "test");
        helper.emit_call(test_num_ser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        // string test
        lget(&mut helper, value_slot, line);
        let test_str_ser = helper.add_import("wasm:js-string", "test");
        helper.emit_call(test_str_ser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        emit_is_array_into(imports, &mut helper, value_slot, line);
        helper.emit_if(line);

        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, out_slot, line);
        push_str(&mut helper, "array", line);
        struct_set_key(&mut helper, SERIAL_KIND_KEY, line);

        helper.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        lset(&mut helper, items_slot, line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let items_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, items_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, value_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        call_ref(&mut helper, 1, line);
        call_import_into(imports, &mut helper, "ecma:array", "push", 2, line);
        helper.emit_op(Op::DROP, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, items_loop, line);
        set_struct_from_slot(&mut helper, out_slot, "items", items_slot, line);

        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "vybe$assoc_keys_csv", line);
        lset(&mut helper, tmp_slot, line);
        lget(&mut helper, tmp_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        helper.emit_else(line);
        { let undef_idx = helper.add_import("wasm:js-undefined", "test"); helper.emit_call(undef_idx, 1, line); }
        helper.emit_if(line);
        helper.emit_else(line);
        lget(&mut helper, tmp_slot, line);
        push_str(&mut helper, "\x1F", line);
        call_import_into(imports, &mut helper, "ecma:string", "split", 2, line);
        lset(&mut helper, names_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let assoc_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        ref_func(&mut helper, helper_idx, line);
        dynamic_get_from_slots(&mut helper, value_slot, key_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, assoc_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, assoc_loop, line);
        set_struct_from_slot(&mut helper, out_slot, "assoc", assoc_slot, line);
        helper.emit_end(line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);

        helper.emit_end(line);

        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__serialize", line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_ser = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_ser, line);
            helper.emit_op(Op::DROP, line);
            lget(&mut helper, fn_slot_ser, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_ser, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_ser, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_ser, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, out_slot, line);
        push_str(&mut helper, "custom_object", line);
        struct_set_key(&mut helper, SERIAL_KIND_KEY, line);
        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__type", line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "class", tmp_slot, line);
        lget(&mut helper, method_slot, line);
        lget(&mut helper, value_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, tmp_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "payload", tmp_slot, line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__sleep", line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_slp = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_slp, line);
            helper.emit_op(Op::DROP, line);
            lget(&mut helper, fn_slot_slp, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_slp, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_slp, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_slp, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, out_slot, line);
        push_str(&mut helper, "sleep_object", line);
        struct_set_key(&mut helper, SERIAL_KIND_KEY, line);
        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__type", line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "class", tmp_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, method_slot, line);
        lget(&mut helper, value_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, names_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let sleep_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        dynamic_get_from_slots(&mut helper, value_slot, key_slot, line);
        lset(&mut helper, tmp_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, tmp_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, assoc_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, sleep_loop, line);
        set_struct_from_slot(&mut helper, out_slot, "fields", assoc_slot, line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, out_slot, line);
        push_str(&mut helper, "object", line);
        struct_set_key(&mut helper, SERIAL_KIND_KEY, line);
        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__type", line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "class", tmp_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, value_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "keys", 1, line);
        lset(&mut helper, names_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let object_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);

        for internal_key in [
            "__type",
            "__types",
            "__control_name",
            "__super",
            "vybe$assoc_keys_csv",
        ] {
            lget(&mut helper, key_slot, line);
            push_str(&mut helper, internal_key, line);
            crate::emitter::ops::emit_dyn_eq(&mut helper, line);
            helper.emit_if(line);
            bump_loop_index(&mut helper, i_slot, line);
            helper.emit_br(1, line);
            helper.emit_end(line);
        }

        dynamic_get_from_slots(&mut helper, value_slot, key_slot, line);
        lset(&mut helper, tmp_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_tmp = alloc_local(&mut helper);
            lget(&mut helper, tmp_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_tmp, line);
            helper.emit_op(Op::DROP, line);
            lget(&mut helper, fn_slot_tmp, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_tmp, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_tmp, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_tmp, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        bump_loop_index(&mut helper, i_slot, line);
        helper.emit_br(1, line);
        helper.emit_end(line);

        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, tmp_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, assoc_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, object_loop, line);
        set_struct_from_slot(&mut helper, out_slot, "fields", assoc_slot, line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
    }

    chunks.push(helper);
    helper_idx
}

fn build_php_unserialize_helper(chunks: &mut Vec<Chunk>, alloc_idx: usize, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut helper = create_function_chunk("__php_unserialize_value", 1);
    helper.local_count = 1;

    let node_slot = 0;
    let _type_slot = alloc_local(&mut helper);
    let kind_slot = alloc_local(&mut helper);
    let out_slot = alloc_local(&mut helper);
    let items_slot = alloc_local(&mut helper);
    let assoc_slot = alloc_local(&mut helper);
    let fields_slot = alloc_local(&mut helper);
    let names_slot = alloc_local(&mut helper);
    let key_slot = alloc_local(&mut helper);
    let tmp_slot = alloc_local(&mut helper);
    let i_slot = alloc_local(&mut helper);
    let n_slot = alloc_local(&mut helper);
    let method_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        emit_nullish_return(&mut helper, node_slot, line);

        // boolean test
        lget(&mut helper, node_slot, line);
        let test_bool_unser = helper.add_import("wasm:js-boolean", "test");
        helper.emit_call(test_bool_unser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        // number test
        lget(&mut helper, node_slot, line);
        let test_num_unser = helper.add_import("wasm:js-number", "test");
        helper.emit_call(test_num_unser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        // string test
        lget(&mut helper, node_slot, line);
        let test_str_unser = helper.add_import("wasm:js-string", "test");
        helper.emit_call(test_str_unser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, SERIAL_KIND_KEY, line);
        lset(&mut helper, kind_slot, line);
        lget(&mut helper, kind_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_else(line);
        { let undef_idx = helper.add_import("wasm:js-undefined", "test"); helper.emit_call(undef_idx, 1, line); }
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_else(line);

        lget(&mut helper, kind_slot, line);
        push_str(&mut helper, "array", line);
        crate::emitter::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        helper.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "items", line);
        lset(&mut helper, items_slot, line);
        lget(&mut helper, items_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let items_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, items_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        lget(&mut helper, out_slot, line);
        lget(&mut helper, tmp_slot, line);
        call_import_into(imports, &mut helper, "ecma:array", "push", 2, line);
        helper.emit_op(Op::DROP, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, items_loop, line);

        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "assoc", line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, assoc_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        helper.emit_else(line);
        { let undef_idx = helper.add_import("wasm:js-undefined", "test"); helper.emit_call(undef_idx, 1, line); }
        helper.emit_if(line);
        helper.emit_else(line);
        lget(&mut helper, assoc_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "keys", 1, line);
        lset(&mut helper, names_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let assoc_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        ref_func(&mut helper, helper_idx, line);
        dynamic_get_from_slots(&mut helper, assoc_slot, key_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, out_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, assoc_loop, line);
        lget(&mut helper, names_slot, line);
        push_str(&mut helper, "\x1F", line);
        call_import_into(imports, &mut helper, "ecma:array", "join", 2, line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "vybe$assoc_keys_csv", tmp_slot, line);
        helper.emit_end(line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);

        helper.emit_end(line);

        ref_func(&mut helper, alloc_idx, line);
        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "class", line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, out_slot, line);

        lget(&mut helper, kind_slot, line);
        push_str(&mut helper, "custom_object", line);
        crate::emitter::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        lget(&mut helper, out_slot, line);
        struct_get_key(&mut helper, "__unserialize", line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_uns = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_uns, line);
            helper.emit_op(Op::DROP, line);
            lget(&mut helper, fn_slot_uns, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_uns, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_uns, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_uns, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        lget(&mut helper, method_slot, line);
        lget(&mut helper, out_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "payload", line);
        call_ref(&mut helper, 1, line);
        call_ref(&mut helper, 2, line);
        helper.emit_op(Op::DROP, line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "fields", line);
        lset(&mut helper, fields_slot, line);
        lget(&mut helper, fields_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "keys", 1, line);
        lset(&mut helper, names_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let fields_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        ref_func(&mut helper, helper_idx, line);
        dynamic_get_from_slots(&mut helper, fields_slot, key_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, out_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, fields_loop, line);

        lget(&mut helper, kind_slot, line);
        push_str(&mut helper, "sleep_object", line);
        crate::emitter::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        lget(&mut helper, out_slot, line);
        struct_get_key(&mut helper, "__wakeup", line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_wk = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_wk, line);
            helper.emit_op(Op::DROP, line);
            lget(&mut helper, fn_slot_wk, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_wk, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_wk, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_wk, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        lget(&mut helper, method_slot, line);
        lget(&mut helper, out_slot, line);
        call_ref(&mut helper, 1, line);
        helper.emit_op(Op::DROP, line);
        helper.emit_end(line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);

        helper.emit_end(line);
        helper.emit_end(line);
    }

    chunks.push(helper);
    helper_idx
}

pub fn emit_php_empty(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let _type_slot = alloc_local(chunk);

    lset(chunk, value_slot, line);

    lget(chunk, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    { let undef_idx = chunk.add_import("wasm:js-undefined", "test"); chunk.emit_call(undef_idx, 1, line); }
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    // boolean test
    lget(chunk, value_slot, line);
    let test_bool_empty = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_empty, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_else(line);

    // number test
    lget(chunk, value_slot, line);
    let test_num_empty = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(test_num_empty, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);

    // string test
    lget(chunk, value_slot, line);
    let test_str_empty = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_empty, 1, line);
    chunk.emit_if_value(line);

    lget(chunk, value_slot, line);
    push_str(chunk, "", line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    push_str(chunk, "0", line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    // array test via ecma:array.isArray
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    let base_len_slot = alloc_local(chunk);
    let extra_len_slot = alloc_local(chunk);
    lset(chunk, base_len_slot, line);

    lget(chunk, value_slot, line);
    let assoc_key = chunk.add_constant(Value::String(Arc::from("vybe$assoc_keys_csv")));
    chunk.emit_op_u16(Op::STRUCT_GET, assoc_key, line);
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, base_len_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);
    push_str(chunk, "\x1F", line);
    let _ = chunk;
    call_import(chunks, current, "ecma:string", "split", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, extra_len_slot, line);
    lget(chunk, base_len_slot, line);
    lget(chunk, extra_len_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_php_serialize(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let helper_idx = build_php_serialize_helper(chunks, line);
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    lset(chunk, value_slot, line);
    ref_func(chunk, helper_idx, line);
    lget(chunk, value_slot, line);
    call_ref(chunk, 1, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:json", "stringify", 1, line);
}

pub fn emit_php_unserialize(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let alloc_idx = build_php_alloc_helper(chunks, line);
    let helper_idx = build_php_unserialize_helper(chunks, alloc_idx, line);
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    lset(chunk, value_slot, line);
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:json", "parse", 1, line);
    let chunk = &mut chunks[current];
    let parsed_slot = alloc_local(chunk);
    lset(chunk, parsed_slot, line);
    ref_func(chunk, helper_idx, line);
    lget(chunk, parsed_slot, line);
    call_ref(chunk, 1, line);
}

/// `WeakReference::create($obj)` → struct with `get` method that derefs the weak ref.
pub fn emit_weak_ref_create(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    // Build the `get()` method: deref the weak ref stored in this.__weak
    let get_idx = {
        let mut c = Chunk::new("__weak_ref_get");
        c.arity = 1; // this
        let weak_k = c.add_constant(Value::String(Arc::from("__weak")));
        c.emit_op_u16(Op::LOCAL_GET, 0, line);
        c.emit_op_u16(Op::STRUCT_GET, weak_k, line);
        { let idx = c.add_import("ecma:weakref", "deref"); c.emit_call(idx, 1, line); }
        c.emit_op(Op::RETURN, line);
        c.local_count = c.local_count.max(1);
        chunks.push(c);
        chunks.len() - 1
    };

    let chunk = &mut chunks[current];
    let obj_slot = chunk.local_count;
    let this_slot = chunk.local_count + 1;
    chunk.local_count += 2;

    // Pop the object arg
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op(Op::DROP, line);

    // Create struct, store weak ref
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op(Op::DROP, line);

    // this.__weak = REF_MAKE_WEAK(obj)
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    { let idx = chunk.add_import("ecma:weakref", "new"); chunk.emit_call(idx, 1, line); }
    let weak_k = chunk.add_constant(Value::String(Arc::from("__weak")));
    chunk.emit_op_u16(Op::STRUCT_SET, weak_k, line);
    chunk.emit_op(Op::DROP, line);

    // Bind get method
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, get_idx as u16, line);
    chunk.emit(0, line);
    let get_k = chunk.add_constant(Value::String(Arc::from("get")));
    chunk.emit_op_u16(Op::STRUCT_SET, get_k, line);
    chunk.emit_op(Op::DROP, line);

    // Return the struct
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}
