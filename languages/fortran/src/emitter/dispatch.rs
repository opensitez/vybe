//! Auto-extracted `fortran.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "fortran.matmul" => {
            crate::emitter::math_adapter::emit_fortran_matmul(chunks, current, argc, line)
        }
        "fortran.max" => {
            crate::emitter::math_adapter::emit_fortran_max(chunks, current, argc, line)
        }
        "fortran.min" => {
            crate::emitter::math_adapter::emit_fortran_min(chunks, current, argc, line)
        }
        "fortran.len_trim" => {
            crate::emitter::string_adapter::emit_fortran_len_trim(chunks, current, line)
        }
        "fortran.adjustl" => {
            crate::emitter::string_adapter::emit_fortran_adjustl(chunks, current, line)
        }

        // ── F2008 bit counting — one core WASM opcode each ──────────────────
        "fortran.popcnt" => {
            crate::emitter::bit_adapter::emit_fortran_popcnt(chunks, current, argc, line)
        }
        "fortran.leadz" => {
            crate::emitter::bit_adapter::emit_fortran_leadz(chunks, current, argc, line)
        }
        "fortran.trailz" => {
            crate::emitter::bit_adapter::emit_fortran_trailz(chunks, current, argc, line)
        }
        "fortran.parity" => {
            crate::emitter::bit_adapter::emit_fortran_parity(chunks, current, argc, line)
        }

        // ── Dart string surfaces (isEmpty / isNotEmpty / replaceFirst etc.) ──
        _ => return false,
    }
    true
}
