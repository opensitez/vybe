//! WASM-compliant opcode definitions.
//!
//! `Op` is a `u32` newtype: `(prefix << 16) | sub_opcode`.
//! Sub-opcode is u16, allowing spec values up to 65535 (relaxed SIMD
//! uses sub-values 256+).
//!
//! Prefixes match the WASM spec:
//! - `0x00`: Core MVP
//! - `0xFB`: GC proposal
//! - `0xFC`: Misc proposal
//! - `0xFD`: SIMD proposal (including relaxed SIMD at sub 256+)
//! - `0xFE`: Threads proposal
//! - `0xF0`: Component Model canonical built-ins (CM3 Binary.md §Canon Definitions)
//! - `0xFF`: VM-internal (not WASM — being eliminated)
//!
//! Opcodes are defined in category files (core.rs, gc.rs, etc.) as `pub const` values.
//! Adding an opcode = one line in one file.

mod canon;
mod core_ops;
mod gc;
mod misc;
pub mod relaxed_simd;
mod simd;
mod threads;
mod vm_internal;

/// A bytecode opcode. Encoded as `(group << 16) | sub_opcode`.
/// Both group and sub are u16. Bytecode stream: [group_hi, group_lo, sub_hi, sub_lo] = 4 bytes.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Op(pub u32);

impl Op {
    /// Create from group (u16) + sub-opcode (u16).
    #[inline]
    pub const fn new(group: u16, sub: u16) -> Self {
        Op(((group as u32) << 16) | sub as u32)
    }

    /// Group (u16). Maps to WASM prefix bytes: 0x00=core, 0xF0=canon, 0xFB=GC, etc.
    #[inline]
    pub const fn group(self) -> u16 {
        (self.0 >> 16) as u16
    }

    /// Sub-opcode within the group (u16).
    #[inline]
    pub const fn sub(self) -> u16 {
        (self.0 & 0xFFFF) as u16
    }

    /// Encode to 4 bytes: [group_hi, group_lo, sub_hi, sub_lo].
    #[inline]
    pub const fn encode(self) -> [u8; 4] {
        let g = self.group();
        let s = self.sub();
        [
            (g >> 8) as u8,
            (g & 0xFF) as u8,
            (s >> 8) as u8,
            (s & 0xFF) as u8,
        ]
    }

    /// Byte length in the internal bytecode stream (always 4).
    #[inline]
    pub const fn encoded_len(self) -> usize {
        4
    }

    /// True if this is a VM-internal opcode (0xFF group), not standard WASM.
    #[inline]
    pub const fn is_vm_internal(self) -> bool {
        self.group() == 0xFF
    }

    /// Decode from group + sub into a validated opcode.
    pub fn decode(group: u16, sub: u16) -> Option<Op> {
        let op = Op::new(group as u16, sub as u16);
        if op.wasm_name_opt().is_some() {
            Some(op)
        } else {
            None
        }
    }

    /// Operand format for this opcode.
    pub fn operand_format(self) -> OperandFormat {
        match self.group() {
            0x00 => core_ops::operand_format(self.sub()),
            0xFB => gc::operand_format(self.sub()),
            0xFC => misc::operand_format(self.sub()),
            0xFD if self.sub() >= 256 => relaxed_simd::operand_format(self.sub()),
            0xFD => simd::operand_format(self.sub()),
            0xFE => threads::operand_format(self.sub()),
            0xF0 => canon::operand_format(self.sub()),
            0xFF => vm_internal::operand_format(self.sub()),
            _ => OperandFormat::None,
        }
    }

    /// WASM disassembly name, or None if not a valid opcode.
    pub fn wasm_name_opt(self) -> Option<&'static str> {
        match self.group() {
            0x00 => core_ops::name(self.sub()),
            0xFB => gc::name(self.sub()),
            0xFC => misc::name(self.sub()),
            0xFD if self.sub() >= 256 => relaxed_simd::name(self.sub()),
            0xFD => simd::name(self.sub()),
            0xFE => threads::name(self.sub()),
            0xF0 => canon::name(self.sub()),
            0xFF => vm_internal::name(self.sub()),
            _ => None,
        }
    }

    /// WASM disassembly name, or "unknown" if not valid.
    pub fn wasm_name(self) -> &'static str {
        self.wasm_name_opt().unwrap_or("unknown")
    }
}

impl std::fmt::Debug for Op {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.wasm_name())
    }
}

impl std::fmt::Display for Op {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.wasm_name())
    }
}

/// Operand format — what follows the 2-byte opcode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub enum OperandFormat {
    /// No operands.
    None,
    /// 1 byte (arg count, upvalue index, lane index).
    U8,
    /// 2 independent bytes.
    U8_U8,
    /// 2 bytes big-endian (const index, local/global index, field name).
    U16,
    /// 2 bytes signed (branch offset).
    I16,
    /// 3 bytes: u16 + u8 (call_import: import_idx + argc).
    U16_U8,
    /// 4 bytes: u16 + u16 (try_start: catch + finally offsets).
    U16_U16,
    /// 4 bytes: u16 + i16 (br_on_cast: type index + branch offset).
    U16_I16,
    /// Unsigned LEB128 u32.
    U32Leb,
    /// Two unsigned LEB128 u32 immediates.
    U32Leb_U32Leb,
    /// WASM memarg: alignment LEB + offset LEB, with an optional memory index
    /// extension when the high bit used by this VM's multi-memory encoding is set.
    MemArg,
    /// WASM memory64 memarg: alignment LEB + u64 offset LEB, with the same
    /// optional memory-index extension as MemArg.
    MemArg64,
    /// Variable: u16 func_idx + u8 upvalue_count + descriptors.
    Closure,
    /// Variable: u32 LEB count + count × u32 LEB labels + u32 LEB default.
    BrTable,
    /// Variable: u8 handler_count + handlers.
    TryTable,
    /// 16 bytes immediate (v128.const).
    V128Const,
    /// 16 bytes lane indices (i8x16.shuffle).
    Shuffle,
    /// Signed LEB128 i32 immediate (i32.const).
    SlI32,
    /// Signed LEB128 i64 immediate (i64.const).
    SlI64,
    /// 4 raw bytes (f32.const).
    RawF32,
    /// 8 raw bytes (f64.const).
    RawF64,
}

impl OperandFormat {
    /// Fixed operand size in bytes, or 0 for variable-length formats.
    pub const fn fixed_size(self) -> usize {
        match self {
            Self::None => 0,
            Self::U8 => 1,
            Self::U8_U8 | Self::U16 | Self::I16 => 2,
            Self::U16_U8 => 3,
            Self::U16_U16 | Self::U16_I16 => 4,
            Self::RawF32 => 4,
            Self::RawF64 => 8,
            Self::V128Const | Self::Shuffle => 16,
            Self::Closure
            | Self::U32Leb
            | Self::U32Leb_U32Leb
            | Self::MemArg
            | Self::MemArg64
            | Self::BrTable
            | Self::TryTable
            | Self::SlI32
            | Self::SlI64 => 0,
        }
    }

    /// Operand size for formats whose size depends on operand bytes.
    pub fn size_in(self, code: &[u8], operand_start: usize) -> usize {
        match self {
            Self::U32Leb => leb_u32_size(code, operand_start),
            Self::U32Leb_U32Leb => {
                let first = leb_u32_size(code, operand_start);
                first + leb_u32_size(code, operand_start + first)
            }
            Self::MemArg => memarg_size(code, operand_start),
            Self::MemArg64 => memarg64_size(code, operand_start),
            Self::Closure => {
                let uv_count_pos = operand_start + 2;
                let uv_count = code.get(uv_count_pos).copied().unwrap_or(0) as usize;
                // u16 func_idx + u8 count + per-upvalue (u8 is_local + u16 index)
                2 + 1 + uv_count * 3
            }
            Self::SlI32 => leb_i32_size(code, operand_start),
            Self::SlI64 => leb_i64_size(code, operand_start),
            Self::BrTable => br_table_size(code, operand_start),
            Self::TryTable => {
                // u8 clause_count + per clause (u8 kind + u16 tag + u16 offset)
                let count = code.get(operand_start).copied().unwrap_or(0) as usize;
                1 + count * 5
            }
            fmt => fmt.fixed_size(),
        }
    }
}

pub fn memarg_size(code: &[u8], operand_start: usize) -> usize {
    let mut ip = operand_start;
    let align_start = ip;
    let align = read_leb_u32(code, &mut ip);
    if ip == align_start {
        return 0;
    }
    let offset_start = ip;
    let _offset = read_leb_u32(code, &mut ip);
    if ip == offset_start {
        return ip.saturating_sub(operand_start);
    }
    if align & 0x40 != 0 {
        let _memidx = read_leb_u32(code, &mut ip);
    }
    ip.saturating_sub(operand_start)
}

pub fn memarg64_size(code: &[u8], operand_start: usize) -> usize {
    let mut ip = operand_start;
    let align_start = ip;
    let align = read_leb_u32(code, &mut ip);
    if ip == align_start {
        return 0;
    }
    let offset_start = ip;
    let _offset = read_leb_u64(code, &mut ip);
    if ip == offset_start {
        return ip.saturating_sub(operand_start);
    }
    if align & 0x40 != 0 {
        let _memidx = read_leb_u32(code, &mut ip);
    }
    ip.saturating_sub(operand_start)
}

pub fn read_leb_u64(code: &[u8], ip: &mut usize) -> u64 {
    let mut result = 0u64;
    let mut shift = 0u32;
    while *ip < code.len() && shift < 64 {
        let byte = code[*ip];
        *ip += 1;
        result |= ((byte & 0x7f) as u64) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    result
}

pub fn leb_i32_size(code: &[u8], start: usize) -> usize {
    leb_u32_size(code, start)
}

pub fn leb_i64_size(code: &[u8], start: usize) -> usize {
    let mut len = 0usize;
    while let Some(byte) = code.get(start + len) {
        len += 1;
        if byte & 0x80 == 0 {
            break;
        }
        if len >= 10 {
            break;
        }
    }
    len
}

pub fn leb_u32_size(code: &[u8], start: usize) -> usize {
    let mut len = 0usize;
    while let Some(byte) = code.get(start + len) {
        len += 1;
        if byte & 0x80 == 0 {
            break;
        }
        if len >= 5 {
            break;
        }
    }
    len
}

pub fn read_leb_u32(code: &[u8], ip: &mut usize) -> u32 {
    let mut result = 0u32;
    let mut shift = 0u32;
    loop {
        let byte = code.get(*ip).copied().unwrap_or(0);
        *ip += 1;
        result |= ((byte & 0x7f) as u32) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
        if shift >= 35 {
            break;
        }
    }
    result
}

pub fn read_leb_i32(code: &[u8], ip: &mut usize) -> i32 {
    let mut result = 0u32;
    let mut shift = 0u32;
    let mut byte;
    loop {
        byte = code.get(*ip).copied().unwrap_or(0);
        *ip += 1;
        result |= ((byte & 0x7f) as u32) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            break;
        }
        if shift >= 35 {
            break;
        }
    }
    if shift < 32 && (byte & 0x40) != 0 {
        result |= !0u32 << shift;
    }
    result as i32
}

pub fn read_leb_i64(code: &[u8], ip: &mut usize) -> i64 {
    let mut result = 0u64;
    let mut shift = 0u32;
    let mut byte;
    loop {
        byte = code.get(*ip).copied().unwrap_or(0);
        *ip += 1;
        result |= ((byte & 0x7f) as u64) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            break;
        }
        if shift >= 70 {
            break;
        }
    }
    if shift < 64 && (byte & 0x40) != 0 {
        result |= !0u64 << shift;
    }
    result as i64
}

pub fn br_table_size(code: &[u8], operand_start: usize) -> usize {
    let mut ip = operand_start;
    let count = read_leb_u32(code, &mut ip) as usize;
    for _ in 0..count {
        let _ = read_leb_u32(code, &mut ip);
    }
    let _default = read_leb_u32(code, &mut ip);
    ip.saturating_sub(operand_start)
}

/// Helper macro for category files. Generates name() and operand_format() functions
/// from a table of [sub] name => format entries.
macro_rules! opcode_category {
    ( $( [$sub:literal] $name:ident => $fmt:ident, $wasm_name:literal; )* ) => {
        pub(super) fn name(sub: u16) -> Option<&'static str> {
            match sub {
                $( $sub => Some($wasm_name), )*
                _ => None,
            }
        }

        pub(super) fn operand_format(sub: u16) -> super::OperandFormat {
            match sub {
                $( $sub => super::OperandFormat::$fmt, )*
                _ => super::OperandFormat::None,
            }
        }
    };
}

// Re-export the macro for category files
pub(crate) use opcode_category;
