use vybe_runtime::{Chunk, Op};

use vybe_compiler::primitives::{callable, collections, strings};

/// Receiver-polymorphic member method: the libc walker/runtime helpers call
/// `indexOf`/`lastIndexOf`/`slice` on BOTH host strings (format parsing,
/// fopen modes) and char arrays (`memrchr` reverse scan). Branch at runtime
/// on `ecma:array.isArray(receiver)` and route to the matching host surface.
/// Stack on entry: receiver, then `argc` args (value-method convention).
fn emit_string_or_array_method(
    chunks: &mut Vec<Chunk>,
    current: usize,
    func: &str,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(1 + argc as u16);
    for slot in (0..=argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    let is_array = chunk.add_import("ecma:array", "isArray");
    chunk.emit_call(is_array, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    chunk.emit_if_value(line);
    for slot in 0..=argc as u16 {
        chunk.emit_op_u16(Op::LOCAL_GET, base + slot, line);
    }
    let arr_fn = chunk.add_import("ecma:array", func);
    chunk.emit_call(arr_fn, argc + 1, line);
    chunk.emit_else(line);
    for slot in 0..=argc as u16 {
        chunk.emit_op_u16(Op::LOCAL_GET, base + slot, line);
    }
    let str_fn = chunk.add_import("ecma:string", func);
    chunk.emit_call(str_fn, argc + 1, line);
    chunk.emit_end(line);
}

/// Stack on entry: count, factory.
/// Returns: an array of `count` fresh values produced by `factory()`.
fn emit_struct_array(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let factory = chunks[current].alloc_scratch(1);
    let count = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, factory, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    collections::emit_new_with_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);

    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, factory, line);
    callable::emit_stacked_invoke(chunks, current, 0, line);
    chunks[current].emit_op(Op::ARRAY_SET, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) -> bool {
    if name.starts_with("libc.") {
        return vybe_platform_libc::emitter::dispatch::dispatch(name, chunks, current, _argc, line);
    }

    match name {
        "c.array_set" => {
            let value = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
            chunks[current].emit_op(Op::ARRAY_SET, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
        }
        "c.index_of" => emit_string_or_array_method(chunks, current, "indexOf", _argc - 1, line),
        "c.last_index_of" => {
            emit_string_or_array_method(chunks, current, "lastIndexOf", _argc - 1, line)
        }
        "c.slice" => emit_string_or_array_method(chunks, current, "slice", _argc - 1, line),
        "c.static_numeric_array" | "c.static_json_value" => {
            let parse = chunks[current].add_import("ecma:json", "parse");
            chunks[current].emit_call(parse, 1, line);
        }
        "c.putchar" => {
            let idx = chunks[current].add_import("wasm:js-string", "fromCharCode");
            chunks[current].emit_call(idx, 1, line);
        }
        // NUL-terminated, not the JS string's length — see
        // `strings::emit_cstr_length`. `wasm:js-string.length` counted past a
        // `'\0'` written into a buffer, so `strlen` disagreed with `cc`.
        "c.strlen" => strings::emit_cstr_length(chunks, current, line),
        "c.strupr" => strings::emit_to_upper(&mut chunks[current], line),
        "c.strlwr" => strings::emit_to_lower(&mut chunks[current], line),
        // `strcmp` compares up to the NUL on BOTH sides; `memcmp` deliberately
        // does not — it is byte-wise over a given length and a NUL is ordinary
        // content. Sharing one binding made `strcmp` compare whole JS strings,
        // so a truncated buffer still compared as its untruncated self.
        "c.strcmp" | "c.strncmp" => {
            let rhs = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, rhs, line);
            strings::emit_cstr_truncate(chunks, current, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, rhs, line);
            strings::emit_cstr_truncate(chunks, current, line);
            let idx = chunks[current].add_import("wasm:js-string", "compare");
            chunks[current].emit_call(idx, 2, line);
        }
        "c.memcmp" => {
            let idx = chunks[current].add_import("wasm:js-string", "compare");
            chunks[current].emit_call(idx, 2, line);
        }
        "c.atoi" | "c.atol" => {
            let idx = chunks[current].add_import("ecma:number", "parseInt");
            chunks[current].emit_call(idx, 1, line);
        }
        "c.qsort" => collections::emit_sort(chunks, current, line),
        "c.struct_array" => emit_struct_array(chunks, current, line),
        _ => return false,
    }
    true
}
