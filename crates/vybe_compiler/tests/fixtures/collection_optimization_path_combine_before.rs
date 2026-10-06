//! Frozen current exact Path.Combine/root predicate emitters.
use vybe_compiler::primitives::{ops, strings};
use vybe_runtime::{Chunk, Op};
const SEP: &str = "/";
pub fn emit_is_rooted(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    // A leading separator of either flavour. A drive letter (`C:`) is a colon
    // at index 1, which is the same test `.NET` makes.
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("/", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("\\", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const(":", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_OR, line);
    ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_combine(chunk: &mut Chunk, argc: u8, line: u32) {
    let n = argc as u16;
    let base = chunk.alloc_scratch(n.max(1));
    for offset in (0..n).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    let acc = chunk.alloc_scratch(1);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    for i in 0..n {
        let seg = base + i;
        // Skip empty segments outright — `Combine("a", "", "b")` is `a/b`.
        chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
        strings::emit_length(chunk, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op(Op::F64_GT, line);
        chunk.emit_if(line);
        {
            chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
            emit_is_rooted(chunk, line);
            ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
            strings::emit_length(chunk, line);
            chunk.emit_f64_const(0.0, line);
            chunk.emit_op(Op::F64_EQ, line);
            chunk.emit_op(Op::I32_OR, line);
            chunk.emit_if_value(line);
            // rooted, or nothing accumulated yet: this segment replaces
            chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
            chunk.emit_else(line);
            {
                // The separator goes in only if the accumulator does not
                // already end with one. `Path.Combine("/a/b/", "c")` is
                // `/a/b/c` in .NET, and `os.path.join`/`path.join` agree — no
                // consumer wants the doubled slash, so this is a bug rather
                // than a dialect. It was visible as `/var/.../T//name`, which
                // still resolves on POSIX and so hid until a test compared
                // the STRING.
                chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
                // "ends with SEP" as `lastIndexOf(SEP) == length - 1`, built
                // from helpers this module already uses. The empty-accumulator
                // case cannot arrive here — the branch above claims it — which
                // matters, because for `""` both sides are -1 and this would
                // answer "already separated".
                chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
                chunk.emit_string_const(SEP, line);
                strings::emit_last_index_of(chunk, line);
                chunk.emit_f64_const(1.0, line);
                chunk.emit_op(Op::F64_ADD, line);
                chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
                strings::emit_length(chunk, line);
                chunk.emit_op(Op::F64_EQ, line);
                chunk.emit_if_value(line);
                chunk.emit_string_const("", line);
                chunk.emit_else(line);
                chunk.emit_string_const(SEP, line);
                chunk.emit_end(line);
                chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
                strings::emit_concat(chunk, 3, line);
            }
            chunk.emit_end(line);
            chunk.emit_op_u16(Op::LOCAL_SET, acc, line);
        }
        chunk.emit_end(line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
}
