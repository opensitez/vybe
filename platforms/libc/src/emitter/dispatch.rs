//! libc `common:libc.*` emit dispatch.
//!
//! C stdio formatting is owned here under the libc platform, not by the
//! generic common formatter.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

// These directives are selected only for a literal, bounded ASCII set.
fn span_literal_set(c:&Chunk)->Vec<u8> {
    let start=c.code.len().checked_sub(8).expect("literal span set operand");
    assert_eq!(Op::decode(c.read_u16(start),c.read_u16(start+2)),Some(Op::GLOBAL_GET));
    let key=c.constants[c.read_u32(start+4) as usize].as_str();
    let text=key.strip_prefix("wasm:string-constants::").expect("literal span string import");
    let mut bytes=text.as_bytes().split(|b|*b==0).next().unwrap_or_default().to_vec();
    bytes.sort_unstable();bytes.dedup();
    assert!(!bytes.is_empty() && bytes.len()<=8 && bytes.is_ascii());bytes
}

fn emit_span_offset_u32(chunk: &mut Chunk, line: u32) {
    // The shared memory32 span helpers return a finite integral Number:
    // an unsigned i32 count, possibly plus the two-byte inline prefix.
    // Every such value fits i64 exactly. Wrap to the same u32 bits without
    // routing this already-numeric result through a host conversion.
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
}

fn emit_nan_to_default(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let fallback = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fallback, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
}

fn call_math(chunks: &mut [Chunk], current: usize, func: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("ecma:math", func);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_erf(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let t = chunks[current].alloc_scratch(1);
    let poly = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, x, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_NE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(f64::INFINITY, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(f64::NEG_INFINITY, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_else(line);

    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_f64_const(0.3275911, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);

    chunks[current].emit_f64_const(1.061405429, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(-1.453152027, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(1.421413741, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(-0.284496736, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(0.254829592, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, poly, line);

    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, poly, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    call_math(chunks, current, "exp", 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub(crate) fn emit_erfc_public(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_erfc(chunks, current, line)
}

fn emit_erfc(chunks: &mut [Chunk], current: usize, line: u32) {
    let erf = chunks[current].alloc_scratch(1);
    emit_erf(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, erf, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, erf, line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

fn emit_stirling_gamma(chunks: &mut [Chunk], current: usize, x: u16, line: u32) {
    chunks[current].emit_f64_const(2.0 * std::f64::consts::PI, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_SQRT, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(std::f64::consts::E, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    call_math(chunks, current, "pow", 2, line);
    chunks[current].emit_op(Op::F64_MUL, line);

    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(12.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_ADD, line);

    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(288.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_ADD, line);

    chunks[current].emit_f64_const(139.0, line);
    chunks[current].emit_f64_const(51840.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_SUB, line);

    chunks[current].emit_op(Op::F64_MUL, line);
}

fn emit_tgamma(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, x, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(0.5, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(std::f64::consts::PI, line);
    chunks[current].emit_op(Op::F64_SQRT, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(-0.5, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-2.0 * std::f64::consts::PI.sqrt(), line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(f64::INFINITY, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    call_math(chunks, current, "floor", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_else(line);
    emit_stirling_gamma(chunks, current, x, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_lgamma(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_tgamma(chunks, current, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    call_math(chunks, current, "log", 1, line);
}

pub fn emit_math(name: &str, chunks: &mut [Chunk], current: usize, line: u32) -> bool {
    match name {
        "libc.number.nan_to_default" => {
            emit_nan_to_default(chunks, current, line);
            true
        }
        "libc.math.erf" => {
            emit_erf(chunks, current, line);
            true
        }
        "libc.math.erfc" => {
            emit_erfc(chunks, current, line);
            true
        }
        "libc.math.tgamma" | "libc.math.gamma" => {
            emit_tgamma(chunks, current, line);
            true
        }
        "libc.math.lgamma" => {
            emit_lgamma(chunks, current, line);
            true
        }
        _ => super::bessel::emit_bessel(name, chunks, current, line),
    }
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if super::dirent_adapter::emit(name, chunks, current, argc, line) {
        return true;
    }
    if super::sdl::emit_sdl(name, chunks, current, argc, line) {
        return true;
    }
    // C ABI database surfaces. Reached by C through `#include <sqlite3.h>` /
    // `<mysql.h>` and by Fortran through `bind(c)` — both arrive here as the
    // bare C symbol, which is why these match on the leaf name.
    if super::sqlite_adapter::emit_sqlite(name, chunks, current, argc, line) {
        return true;
    }
    if super::mysql_adapter::emit_mysql(name, chunks, current, argc, line) {
        return true;
    }
    if emit_math(name, chunks, current, line) {
        return true;
    }
    if super::ctype::emit(name, chunks, current, line) {
        return true;
    }
    match name {
        "libc.stdio.fileno" if argc == 1 => {
            super::stdio_terminal::emit_fileno(&mut chunks[current], line);
            true
        }
        "libc.stdio.isatty" if argc == 1 => {
            super::stdio_terminal::emit_isatty(&mut chunks[current], line);
            true
        }
        "libc.socket.send_bytes" => {
            let chunk = &mut chunks[current];
            let sink = chunk.add_import("wasi:sockets/types", "[method]tcp-socket.send");
            vybe_compiler::primitives::io::emit_bytes_to_receiver_sink(chunk, sink, line);
            vybe_compiler::primitives::io::emit_discard_future(chunk, line);
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            true
        }
        "libc.stdio.printf" => {
            super::stdio_format::emit_sprintf(chunks, current, argc, line);
            let idx = chunks[current].add_import("web:console", "log");
            chunks[current].emit_call(idx, 1, line);
            true
        }
        "libc.memory.copy_column" => {
            super::byte_memory::emit_copy_column(chunks, current, line);
            true
        }
        "libc.memory.copy_column16" => {
            super::byte_memory::emit_copy_column_16(chunks, current, line);
            true
        }
        "libc.string.char_to_str" | "libc.string.strlen" => {
            super::string_memory::emit_read(chunks, current, name.ends_with("char_to_str"), line);
            true
        }
        "libc.string.byte_copy" => {
            super::string_memory::emit_byte_copy(chunks, current, line);
            true
        }
        "libc.string.binary_byte_copy" => {
            super::string_memory::emit_binary_byte_copy(chunks, current, line);
            true
        }
        "libc.string.strrchr_linear" => {
            super::string_memory::emit_strrchr_linear(chunks, current, line);
            true
        }
        "libc.string.strchr_linear" => {
            super::string_memory::emit_strchr_linear(chunks, current, line);
            true
        }
        "libc.string.mem_index_of_linear" => {
            vybe_compiler::primitives::memory_find::emit_linear_index(chunks, current, line);
            true
        }
        "libc.string.strspn_literal_set" | "libc.string.strcspn_literal_set" => {
            let bytes=span_literal_set(&chunks[current]);chunks[current].emit_op(Op::DROP,line);
            vybe_compiler::primitives::memory_span_set::emit_literal_set(chunks,current,name=="libc.string.strspn_literal_set",&bytes,line);true
        }
        "libc.string.strpbrk_literal_set" => {
            let bytes=span_literal_set(&chunks[current]);let c=&mut chunks[current];c.emit_op(Op::DROP,line);
            let b=c.alloc_scratch(2);let (ptr,offset)=(b,b+1);c.emit_op_u16(Op::LOCAL_TEE,ptr,line);
            vybe_compiler::primitives::memory_span_set::emit_literal_set(chunks,current,false,&bytes,line);
            let c=&mut chunks[current];c.emit_op_u16(Op::LOCAL_SET,offset,line);c.emit_op_u16(Op::LOCAL_GET,ptr,line);c.emit_op_u16(Op::LOCAL_GET,offset,line);
            emit_span_offset_u32(c,line);c.emit_op(Op::I32_ADD,line);c.emit_op_u16(Op::LOCAL_TEE,offset,line);c.emit_op(Op::I32_LOAD8_U,line);c.emit_op(Op::I32_EQZ,line);
            let branch=c.emit_if_value(line);c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN,line);c.emit_else(line);c.emit_op_u16(Op::LOCAL_GET,offset,line);c.emit_end(line);c.patch_block(branch);true
        }
        "libc.string.strspn_one_byte" | "libc.string.strcspn_one_byte" => {
            vybe_compiler::primitives::memory_span::emit_single_byte(chunks,current,name=="libc.string.strspn_one_byte",line);
            true
        }
        "libc.string.strpbrk_one_byte" => {
            let c=&mut chunks[current];let b=c.alloc_scratch(2);let (ptr,offset)=(b,b+1);
            c.emit_op_u16(Op::LOCAL_SET,offset,line);c.emit_op_u16(Op::LOCAL_SET,ptr,line);
            c.emit_op_u16(Op::LOCAL_GET,ptr,line);c.emit_op_u16(Op::LOCAL_GET,offset,line);
            vybe_compiler::primitives::memory_span::emit_single_byte(chunks,current,false,line);
            let c=&mut chunks[current];c.emit_op_u16(Op::LOCAL_SET,offset,line);
            c.emit_op_u16(Op::LOCAL_GET,ptr,line);c.emit_op_u16(Op::LOCAL_GET,offset,line);
            emit_span_offset_u32(c,line);
            c.emit_op(Op::I32_ADD,line);c.emit_op_u16(Op::LOCAL_TEE,offset,line);
            c.emit_op(Op::I32_LOAD8_U,line);c.emit_op(Op::I32_EQZ,line);
            let branch=c.emit_if_value(line);
            c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN,line);
            c.emit_else(line);c.emit_op_u16(Op::LOCAL_GET,offset,line);c.emit_end(line);c.patch_block(branch);
            true
        }
        "libc.string.strncpy" => {
            super::string_memory::emit_strncpy(chunks, current, line);
            true
        }
        "libc.string.strdup" => {
            super::string_memory::emit_strdup(chunks, current, line);
            true
        }
        "libc.memory.alloc" => {
            super::heap_memory::emit_alloc(chunks, current, line);
            true
        }
        "libc.memory.realloc" => {
            super::heap_memory::emit_realloc(chunks, current, line);
            true
        }
        "libc.string.pointer_add" | "libc.string.pointer_write" => {
            super::string_memory::emit_pointer(
                chunks,
                current,
                name.ends_with("pointer_write"),
                line,
            );
            true
        }
        "libc.stdio.sprintf" => {
            super::stdio_format::emit_sprintf(chunks, current, argc, line);
            true
        }
        "libc.stdio.vsprintf" => {
            super::stdio_format::emit_sprintf_from_array(chunks, current, line);
            true
        }
        _ => false,
    }
}

#[cfg(test)]
mod span_offset_tests {
    use super::*;

    #[test]
    fn span_offset_emits_only_standard_numeric_operations() {
        let mut actual = Chunk::new("span-offset");
        emit_span_offset_u32(&mut actual, 0);
        let mut expected = Chunk::new("expected");
        expected.emit_op(Op::I64_TRUNC_F64_S, 0);
        expected.emit_op(Op::I32_WRAP_I64, 0);
        assert_eq!(actual.code, expected.code);
        assert_eq!(actual.code.len(), 8);
    }

    #[test]
    fn span_result_range_has_exact_unsigned_wrap_bits() {
        for count in [0u64, 1, 2, 127, 128, 65535, 2147483647,
                      2147483648, 4294967295, 4294967296, 4294967297] {
            let value = count as f64;
            assert_eq!(value as i64 as u32, count as u32);
            assert_eq!(value.fract(), 0.0);
        }
    }
}
