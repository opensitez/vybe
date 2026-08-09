use vybe_runtime::{Chunk, Op};

use vybe_compiler::primitives::{collections, strings};

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

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) -> bool {
    if name.starts_with("libc.") {
        return vybe_platform_libc::emitter::dispatch::dispatch(name, chunks, current, _argc, line);
    }

    match name {
        "c.index_of" => emit_string_or_array_method(chunks, current, "indexOf", _argc - 1, line),
        "c.last_index_of" => {
            emit_string_or_array_method(chunks, current, "lastIndexOf", _argc - 1, line)
        }
        "c.slice" => emit_string_or_array_method(chunks, current, "slice", _argc - 1, line),
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
        _ => return false,
    }
    true
}
