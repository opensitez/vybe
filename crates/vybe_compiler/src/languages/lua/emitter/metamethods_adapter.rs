//! Lua metamethods runtime support.
//! 
//! Handles metamethod lookups and calls for:
//! - Arithmetic: __add, __sub, __mul, __div, __mod, __pow, __unm, __idiv
//! - Comparison: __lt, __le, __eq
//! - Indexing: __index, __newindex, __len
//! - Other: __concat, __call, __tostring
//!
//! These are emitted as runtime calls, not AST desugaring.

use vybe_bytecode::Chunk;

/// Emit bytecode to call __add metamethod
pub fn emit_metamethod_add(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_add");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_sub(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_sub");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_mul(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_mul");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_div(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_div");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_mod(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_mod");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_pow(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_pow");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_unm(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_unm");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_lt(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_lt");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_le(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_le");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_eq(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_eq");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_index(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_index");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_newindex(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_newindex");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_len(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_len");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_concat(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_concat");
    chunk.emit_call(host_idx, argc, line);
}

pub fn emit_metamethod_call(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let host_idx = chunk.add_import("ecma:lua", "metamethod_call");
    chunk.emit_call(host_idx, argc, line);
}
