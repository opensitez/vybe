//! # nontrapping-float-to-int-conversions proposal
//!
//! Spec: `proposals/nontrapping-float-to-int-conversions/`
//! Adds saturating (non-trapping) float → integer truncation:
//! NaN → 0, overflow → integer min/max instead of an arithmetic trap.
//!
//! All 8 opcodes use prefix `0xFC`, sub-bytes `0x00–0x07`. No immediates.
//!
//! ## Status in Vybe
//!
//! | Op                      | Opcode       | Status |
//! |-------------------------|--------------|--------|
//! | `i32.trunc_sat_f32_s`   | 0xFC 0x00    | ✅ Rust `as` saturates |
//! | `i32.trunc_sat_f32_u`   | 0xFC 0x01    | ✅ |
//! | `i32.trunc_sat_f64_s`   | 0xFC 0x02    | ✅ |
//! | `i32.trunc_sat_f64_u`   | 0xFC 0x03    | ✅ |
//! | `i64.trunc_sat_f32_s`   | 0xFC 0x04    | ✅ |
//! | `i64.trunc_sat_f32_u`   | 0xFC 0x05    | ✅ |
//! | `i64.trunc_sat_f64_s`   | 0xFC 0x06    | ✅ |
//! | `i64.trunc_sat_f64_u`   | 0xFC 0x07    | ✅ |
//!
//! ## Implementation note
//!
//! Rust's `as` cast (f64 → i32/u32/i64/u64) has saturating semantics since
//! Rust 1.45 (stabilised via `saturating_cast` RFC). Specifically:
//! - NaN → 0
//! - value > INT_MAX → INT_MAX (or UINT_MAX for unsigned)
//! - value < INT_MIN → INT_MIN
//! This matches the WASM spec for `trunc_sat` exactly, so no extra branching
//! is needed in the dispatch cases.
//!
//! The dispatch cases live in `dispatch.rs` under the heading
//! "nontrapping-float-to-int-conversions proposal".
//! The reader mapping (0xFC 0x00–0x07 → Op constants) lives in `reader.rs`.
//! The writer pass-through lives in `code.rs` (0xFC `_ =>` case already
//! emits `prefix + leb128(sub)` with no further immediates — correct).

use vybe_runtime::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}
