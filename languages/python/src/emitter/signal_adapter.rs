//! `signal` adapters.
//!
//! WASI/ECMA does not deliver POSIX signals to Python code here, so the Python
//! surface is deliberately deterministic: handlers can be installed and queried
//! as harmless values, alarms do not sleep, and descriptions are printable.

use vybe_runtime::Chunk;

use super::adapter_util::stash_args;

pub fn emit_getsignal(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let _ = stash_args(chunks, current, argc, line);
    chunks[current].emit_f64_const(0.0, line);
}

pub fn emit_signal(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let _ = stash_args(chunks, current, argc, line);
    chunks[current].emit_f64_const(0.0, line);
}

pub fn emit_alarm(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let _ = stash_args(chunks, current, argc, line);
    chunks[current].emit_f64_const(0.0, line);
}

pub fn emit_pause(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let _ = stash_args(chunks, current, argc, line);
    chunks[current].emit_f64_const(0.0, line);
}

pub fn emit_strsignal(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let _ = stash_args(chunks, current, argc, line);
    chunks[current].emit_string_const("Interrupt", line);
}
