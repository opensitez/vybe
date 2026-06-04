//! # simd proposal
//!
//! Spec: `proposals/spec/proposals/simd/`. 128-bit vector
//! instructions. Prefix `0xFD`. The opcodes themselves are declared in
//! `opcode/simd.rs`; the emitter lowers them via the generic
//! "prefix-and-sub" fall-through in `code.rs` (SIMD ops are just
//! pass-through because they have no import or type-context needs).
//!
//! ## Status in Vybe
//!
//! | Family                 | Status |
//! |------------------------|--------|
//! | `v128.load` / `store`  | ✅ all load/store addressing modes |
//! | splat / extract / replace lanes | ✅ i8x16/i16x8/i32x4/i64x2/f32x4/f64x2 |
//! | arithmetic (i/f)       | ✅ add/sub/mul/div/min/max/abs/neg |
//! | comparisons            | ✅ eq/ne/lt/le/gt/ge per lane type |
//! | bitwise                | ✅ and/or/xor/not/andnot/bitselect |
//! | shuffles / swizzle     | ✅ |
//! | relaxed SIMD (0xFD 0x100+) | ⚠  opcode declared; lowering not verified |
//!
//! ## VM side
//!
//! The VM stores SIMD vectors as `Value::V128([u8; 16])`. All SIMD
//! opcodes are implemented in `vm.rs` via the `simd_helpers` group.

use crate::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}
