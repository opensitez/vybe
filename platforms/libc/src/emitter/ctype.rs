//! Inline C-locale byte classification with one evaluation of the argument.

use vybe_runtime::{Chunk, opcode::Op};

fn compare(chunk: &mut Chunk, value: u16, op: Op, bound: i32, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_i32_const(bound, line);
    chunk.emit_op(op, line);
}

fn range(chunk: &mut Chunk, value: u16, lo: i32, hi: i32, line: u32) {
    compare(chunk, value, Op::I32_GE_S, lo, line);
    compare(chunk, value, Op::I32_LE_S, hi, line);
    chunk.emit_op(Op::I32_AND, line);
}

fn alpha(chunk: &mut Chunk, value: u16, line: u32) {
    range(chunk, value, 65, 90, line);
    range(chunk, value, 97, 122, line);
    chunk.emit_op(Op::I32_OR, line);
}

pub fn emit(name: &str, chunks: &mut [Chunk], current: usize, line: u32) -> bool {
    let Some(kind) = name.strip_prefix("libc.ctype.") else {
        return false;
    };
    if !matches!(
        kind,
        "isalpha"
            | "isdigit"
            | "isalnum"
            | "isspace"
            | "isupper"
            | "islower"
            | "isxdigit"
            | "iscntrl"
            | "isprint"
            | "ispunct"
            | "isgraph"
            | "isblank"
            | "toupper"
            | "tolower"
    ) {
        return false;
    }

    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    match kind {
        "isalpha" => alpha(chunk, value, line),
        "isdigit" => range(chunk, value, 48, 57, line),
        "isalnum" => {
            alpha(chunk, value, line);
            range(chunk, value, 48, 57, line);
            chunk.emit_op(Op::I32_OR, line);
        }
        "isspace" => {
            compare(chunk, value, Op::I32_EQ, 32, line);
            range(chunk, value, 9, 13, line);
            chunk.emit_op(Op::I32_OR, line);
        }
        "isupper" => range(chunk, value, 65, 90, line),
        "islower" => range(chunk, value, 97, 122, line),
        "isxdigit" => {
            range(chunk, value, 48, 57, line);
            range(chunk, value, 65, 70, line);
            chunk.emit_op(Op::I32_OR, line);
            range(chunk, value, 97, 102, line);
            chunk.emit_op(Op::I32_OR, line);
        }
        "iscntrl" => {
            range(chunk, value, 0, 31, line);
            compare(chunk, value, Op::I32_EQ, 127, line);
            chunk.emit_op(Op::I32_OR, line);
        }
        "isprint" => range(chunk, value, 32, 126, line),
        "isgraph" => range(chunk, value, 33, 126, line),
        "isblank" => {
            compare(chunk, value, Op::I32_EQ, 32, line);
            compare(chunk, value, Op::I32_EQ, 9, line);
            chunk.emit_op(Op::I32_OR, line);
        }
        "ispunct" => {
            range(chunk, value, 33, 126, line);
            alpha(chunk, value, line);
            range(chunk, value, 48, 57, line);
            chunk.emit_op(Op::I32_OR, line);
            chunk.emit_op(Op::I32_EQZ, line);
            chunk.emit_op(Op::I32_AND, line);
        }
        "toupper" | "tolower" => {
            let upper = kind == "toupper";
            range(
                chunk,
                value,
                if upper { 97 } else { 65 },
                if upper { 122 } else { 90 },
                line,
            );
            chunk.emit_if(line);
            chunk.emit_op_u16(Op::LOCAL_GET, value, line);
            chunk.emit_i32_const(if upper { -32 } else { 32 }, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value, line);
            chunk.emit_end(line);
            chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        }
        _ => unreachable!(),
    }
    true
}
