use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::object::{
    emit_bind_bound_method, emit_bind_getter, emit_bind_setter,
};

const SERIAL_KIND_KEY: &str = "vybe$php_ser_kind";

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
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
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, argc, line);
}

fn ref_func(chunk: &mut Chunk, func_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_idx as u16, line);
    chunk.emit(0, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let cs_slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &cs_slot, Dest::Stack, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let cs_slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_slot, ValueSource::Stack, line);
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
}

fn set_struct_from_slot(chunk: &mut Chunk, obj_slot: u16, key: &str, value_slot: u16, line: u32) {
    lget(chunk, obj_slot, line);
    lget(chunk, value_slot, line);
    struct_set_key(chunk, &ClassSlot::internal(key), line);
}

fn helper_loop_start(chunk: &mut Chunk, line: u32) -> vybe_compiler::primitives::loops::LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    vybe_compiler::primitives::loops::LoopState {
        block_patch,
        loop_patch,
        body_block_patch: None,
    }
}

fn helper_loop_cond(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
}

fn helper_loop_end(
    chunk: &mut Chunk,
    state: vybe_compiler::primitives::loops::LoopState,
    line: u32,
) {
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
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_else(line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    chunk.emit_if(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_is_array_into(imports: &mut Chunk, code: &mut Chunk, value_slot: u16, line: u32) {
    lget(code, value_slot, line);
    call_import_into(imports, code, "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(code, line);
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
    helper.alloc_scratch(1);
    let class_slot = 0;
    let obj_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        emit_nullish_return(&mut helper, class_slot, line);

        for (idx, ty) in types.iter().enumerate().filter(|(_, ty)| !ty.is_interface) {
            lget(&mut helper, class_slot, line);
            push_str(&mut helper, &ty.name, line);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
            helper.emit_if(line);

            // `struct.new_default $T` — `unserialize()` revives an instance
            // without running its constructor, so allocation is the ONLY place
            // its identity can be established, and the spec instruction does it
            // there. `idx + 1` is the 1-based module type-table index, the same
            // convention `reserve_type_slot` / `register_gc_array_type` use.
            helper.emit_op_u16(Op::STRUCT_NEW_DEFAULT, idx as u16 + 1, line);
            lset(&mut helper, obj_slot, line);

            lget(&mut helper, obj_slot, line);
            push_str(&mut helper, &ty.name, line);
            let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
            class_slots::emit_class_set(&mut helper, ObjSource::Stack, &cs_id, ValueSource::Stack, line);

            lget(&mut helper, obj_slot, line);
            push_str(&mut helper, &ty.name.to_lowercase(), line);
            struct_set_key(&mut helper, &ClassSlot::internal("__control_name"), line);

            // No identity stamp here — `struct.new_default $T` above already
            // set the rtt at allocation. The `GLOBAL_GET __tid_<class>` +
            // `__type_id` property write that used to live here fed a property
            // nothing in the VM ever reads.

            for field in &ty.fields {
                lget(&mut helper, obj_slot, line);
                helper.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
                struct_set_key(&mut helper, &ClassSlot::internal(field), line);
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
                    emit_bind_bound_method(
                        &mut helper,
                        obj_slot,
                        method_name,
                        *method_chunk_idx,
                        None,
                        false, // PHP binds the receiver at call time, not on access
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

fn build_php_serialize_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut helper = create_function_chunk("__php_serialize_value", 1);
    helper.alloc_scratch(1);

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
        struct_set_key(&mut helper, &ClassSlot::internal(SERIAL_KIND_KEY), line);

        helper.emit_array_new_fixed(0, 0, line);
        lset(&mut helper, items_slot, line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let items_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
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
        struct_get_key(&mut helper, &ClassSlot::internal("vybe$assoc_keys_csv"), line);
        lset(&mut helper, tmp_slot, line);
        lget(&mut helper, tmp_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        helper.emit_else(line);
        {
            let undef_idx = helper.add_import("wasm:js-undefined", "test");
            helper.emit_call(undef_idx, 1, line);
        }
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
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
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
        struct_get_key(&mut helper, &ClassSlot::internal("__serialize"), line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_ser = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_ser, line);
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
        struct_set_key(&mut helper, &ClassSlot::internal(SERIAL_KIND_KEY), line);
        lget(&mut helper, value_slot, line);
        let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
        class_slots::emit_class_get(&mut helper, ObjSource::Stack, &cs_id, Dest::Stack, line);
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
        struct_get_key(&mut helper, &ClassSlot::internal("__sleep"), line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_slp = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_slp, line);
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
        struct_set_key(&mut helper, &ClassSlot::internal(SERIAL_KIND_KEY), line);
        lget(&mut helper, value_slot, line);
        let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
        class_slots::emit_class_get(&mut helper, ObjSource::Stack, &cs_id, Dest::Stack, line);
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
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
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
        struct_set_key(&mut helper, &ClassSlot::internal(SERIAL_KIND_KEY), line);
        lget(&mut helper, value_slot, line);
        let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
        class_slots::emit_class_get(&mut helper, ObjSource::Stack, &cs_id, Dest::Stack, line);
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
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
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
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
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
    helper.alloc_scratch(1);

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
        struct_get_key(&mut helper, &ClassSlot::internal(SERIAL_KIND_KEY), line);
        lset(&mut helper, kind_slot, line);
        lget(&mut helper, kind_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_else(line);
        {
            let undef_idx = helper.add_import("wasm:js-undefined", "test");
            helper.emit_call(undef_idx, 1, line);
        }
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_else(line);

        lget(&mut helper, kind_slot, line);
        push_str(&mut helper, "array", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        helper.emit_array_new_fixed(0, 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, &ClassSlot::internal("items"), line);
        lset(&mut helper, items_slot, line);
        lget(&mut helper, items_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let items_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
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
        struct_get_key(&mut helper, &ClassSlot::internal("assoc"), line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, assoc_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        helper.emit_else(line);
        {
            let undef_idx = helper.add_import("wasm:js-undefined", "test");
            helper.emit_call(undef_idx, 1, line);
        }
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
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
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
        struct_get_key(&mut helper, &ClassSlot::internal("class"), line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, out_slot, line);

        lget(&mut helper, kind_slot, line);
        push_str(&mut helper, "custom_object", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        lget(&mut helper, out_slot, line);
        struct_get_key(&mut helper, &ClassSlot::internal("__unserialize"), line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_uns = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_uns, line);
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
        struct_get_key(&mut helper, &ClassSlot::internal("payload"), line);
        call_ref(&mut helper, 1, line);
        call_ref(&mut helper, 2, line);
        helper.emit_op(Op::DROP, line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, &ClassSlot::internal("fields"), line);
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
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
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
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        lget(&mut helper, out_slot, line);
        struct_get_key(&mut helper, &ClassSlot::internal("__wakeup"), line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_wk = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_wk, line);
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

pub fn emit_php_unserialize(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let alloc_idx = build_php_alloc_helper(chunks, line);
    let helper_idx = build_php_unserialize_helper(chunks, alloc_idx, line);
    let chunk = &mut chunks[current];
    if argc > 1 {
        let _options_slot = alloc_local(chunk);
        lset(chunk, _options_slot, line);
    }
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
