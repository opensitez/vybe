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

/// Abstract heap types (GC proposal §Reference types) — the single-byte
/// heaptype immediate of `ref.null`, `ref.test`, `ref.cast`, `br_on_cast`.
///
/// These live here, beside the opcodes that take them, because they are
/// instruction immediates: `ref.null` cannot be emitted without one.
/// `platforms/wasm/src/encoding.rs` re-exports them so the binary writer and
/// the bytecode emitter cannot disagree about a byte value.
pub mod heaptype {
    pub const HT_NOFUNC: u8 = 0x73; // -0x0D
    pub const HT_NOEXTERN: u8 = 0x72; // -0x0E
    pub const HT_NONE: u8 = 0x71; // -0x0F (nullref)
    pub const HT_FUNC: u8 = 0x70; // -0x10 (funcref)
    pub const HT_EXTERN: u8 = 0x6F; // -0x11 (externref)
    pub const HT_ANY: u8 = 0x6E; // -0x12 (anyref)
    pub const HT_EQ: u8 = 0x6D; // -0x13 (eqref)
    pub const HT_I31: u8 = 0x6C; // -0x14 (i31ref)
    pub const HT_STRUCT: u8 = 0x6B; // -0x15 (structref)
    pub const HT_ARRAY: u8 = 0x6A; // -0x16 (arrayref)

    /// A heap type, exactly as the spec's immediate encodes one: a single
    /// signed LEB where a NEGATIVE value is one of the abstract types above
    /// and a NON-NEGATIVE value is an index into the module's type section.
    /// There is no third case — in particular there is no name.
    ///
    /// `Concrete` carries the 1-based module type index (`chunk_type_base` +
    /// this - 1 → registry id), the same immediate `struct.new` uses.
    #[derive(Clone, Copy, PartialEq, Eq, Debug)]
    pub enum HeapType {
        /// One of the `HT_*` constants above.
        Abstract(u8),
        /// Module type index.
        Concrete(u32) }

    impl HeapType {
        /// Decode the spec immediate. Negative is abstract — and the `HT_*`
        /// constants ARE the low 7 bits of those negative values, which is why
        /// masking recovers the byte (`-0x12 & 0x7F == 0x6E == HT_ANY`).
        #[inline]
        pub const fn from_sleb(value: i32) -> HeapType {
            if value < 0 {
                HeapType::Abstract((value & 0x7F) as u8)
            } else {
                HeapType::Concrete(value as u32)
            }
        }

        /// Encode as the spec immediate — the inverse of [`HeapType::from_sleb`].
        #[inline]
        pub const fn to_sleb(self) -> i32 {
            match self {
                // Sign-extend the 7-bit abstract encoding back to negative.
                HeapType::Abstract(byte) => (byte as i32) | !0x7F,
                HeapType::Concrete(index) => index as i32 }
        }

        /// The abstract heap type spelled by `name`, if it is one. The spec's
        /// own spelling only — `"function"`, `"object"`, `"number"` and the
        /// rest are LANGUAGE type tags and must not resolve here.
        pub fn from_spec_name(name: &str) -> Option<HeapType> {
            Some(HeapType::Abstract(match name {
                "any" => HT_ANY,
                "eq" => HT_EQ,
                "i31" => HT_I31,
                "struct" => HT_STRUCT,
                "array" => HT_ARRAY,
                "func" => HT_FUNC,
                "extern" => HT_EXTERN,
                "none" => HT_NONE,
                "nofunc" => HT_NOFUNC,
                "noextern" => HT_NOEXTERN,
                _ => return None }))
        }
    }

    /// Is this heaptype in the GC heap? A `ref.null` of a GC type is a WASM GC
    /// **typed null** — the GC accessors (`struct.get`/`set`, `array.*`) trap
    /// on it per spec — whereas `ref.null extern` / `ref.null func` is the
    /// lenient null the dynamic languages use. This predicate is what the two
    /// used to be told apart by, back when the GC case had its own custom
    /// opcode instead of an immediate.
    #[inline]
    pub const fn is_gc_heap(ht: u8) -> bool {
        !matches!(ht, HT_EXTERN | HT_FUNC | HT_NOEXTERN | HT_NOFUNC)
    }
}

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
            _ => OperandFormat::None }
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
            _ => None }
    }

    /// Resolve a WASM mnemonic (e.g. `"i32.clz"`) to its opcode. The inverse of
    /// [`Op::wasm_name_opt`], sourced from the same per-category tables so callers
    /// (like the compiler's `opcode:` emit strategy) never maintain a second list.
    /// Categories are probed core → GC → misc → SIMD → threads → canon; mnemonics
    /// are unique across them.
    pub fn from_wasm_name(wasm_name: &str) -> Option<Op> {
        if let Some(sub) = core_ops::from_name(wasm_name) {
            return Some(Op::new(0x00, sub));
        }
        if let Some(sub) = gc::from_name(wasm_name) {
            return Some(Op::new(0xFB, sub));
        }
        if let Some(sub) = misc::from_name(wasm_name) {
            return Some(Op::new(0xFC, sub));
        }
        if let Some(sub) = simd::from_name(wasm_name) {
            return Some(Op::new(0xFD, sub));
        }
        if let Some(sub) = relaxed_simd::from_name(wasm_name) {
            return Some(Op::new(0xFD, sub));
        }
        if let Some(sub) = threads::from_name(wasm_name) {
            return Some(Op::new(0xFE, sub));
        }
        if let Some(sub) = canon::from_name(wasm_name) {
            return Some(Op::new(0xF0, sub));
        }
        if let Some(sub) = vm_internal::from_name(wasm_name) {
            return Some(Op::new(0xFF, sub));
        }
        None
    }

    /// WASM disassembly name, or "unknown" if not valid.
    pub fn wasm_name(self) -> &'static str {
        self.wasm_name_opt().unwrap_or("unknown")
    }

    /// Like [`from_wasm_name`](Self::from_wasm_name) but tolerant of the wast
    /// walker's name flattening: the walker turns a mnemonic's single `.` into
    /// `_`, and some namespaces already contain `_` (`stringview_wtf16.length`,
    /// `stringview_wtf8.advance`), so the dot's original position is ambiguous.
    /// Try the exact name, then each underscore position as the dot; return the
    /// first that resolves. Only reached when an exact lookup already failed, so
    /// normal (already-dotted) callers are unaffected.
    pub fn from_flattened_name(name: &str) -> Option<Op> {
        if let Some(op) = Self::from_wasm_name(name) {
            return Some(op);
        }
        let positions: Vec<usize> = name.match_indices('_').map(|(i, _)| i).collect();
        for i in positions {
            let mut cand = name.to_string();
            cand.replace_range(i..i + 1, ".");
            if let Some(op) = Self::from_wasm_name(&cand) {
                return Some(op);
            }
        }
        None
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
    /// 3 independent bytes (call_indirect: argc + tableidx + result count).
    U8_U8_U8,
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
    /// OPTIONAL marker-tagged memarg — the shared shape for every memory
    /// access whose memarg may be elided internally: SIMD (`v128.load` /
    /// `v128.store` / splat/zero loads) AND the core loads/stores
    /// (0x28-0x3E). Present iff the first LEB's 0x80 bit is set (0x100 =
    /// memory64 offset, 0x40 = memidx extension follows — the spec
    /// multi-memory bit); a compiler-emitted op with no memarg contributes
    /// ZERO operand bytes (the peek is unambiguous: instruction group-hi
    /// bytes are always 0x00). The spec binary always writes a memarg; the
    /// writer materializes natural align + offset 0 when absent.
    SimdMemArg,
    /// WASM SIMD lane memory op (`v128.load8_lane` / `v128.store32_lane` / …):
    /// the same optional marker-tagged memarg as `SimdMemArg`, followed by a
    /// single lane-index byte (lane < 0x80, so the peek never misfires).
    MemLane,
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
    RawF64 }

impl OperandFormat {
    /// Fixed operand size in bytes, or 0 for variable-length formats.
    pub const fn fixed_size(self) -> usize {
        match self {
            Self::None => 0,
            Self::U8 => 1,
            Self::U8_U8 | Self::U16 | Self::I16 => 2,
            Self::U8_U8_U8 | Self::U16_U8 => 3,
            Self::U16_U16 | Self::U16_I16 => 4,
            Self::RawF32 => 4,
            Self::RawF64 => 8,
            Self::V128Const | Self::Shuffle => 16,
            Self::Closure
            | Self::U32Leb
            | Self::U32Leb_U32Leb
            | Self::MemArg
            | Self::MemArg64
            | Self::SimdMemArg
            | Self::MemLane
            | Self::BrTable
            | Self::TryTable
            | Self::SlI32
            | Self::SlI64 => 0 }
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
            Self::SimdMemArg => simd_memarg_size(code, operand_start),
            // Optional marker-tagged memarg + the mandatory lane byte.
            Self::MemLane => simd_memarg_size(code, operand_start) + 1,
            Self::Closure => {
                let uv_count_pos = operand_start + 2;
                // Mask the 0x80 "no-intern" flag (see REF_FUNC dispatch).
                let uv_count = (code.get(uv_count_pos).copied().unwrap_or(0) & 0x7f) as usize;
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
            fmt => fmt.fixed_size() }
    }
}

/// Size of the OPTIONAL marker-tagged SIMD memarg. Mirrors the dispatch peek
/// (`read_optional_simd_memarg`) exactly: no 0x80 marker on the first LEB →
/// no memarg present → 0 bytes. Present: align LEB + offset LEB (u64 when the
/// 0x100 flag is set — LEBs self-delimit, so the byte count is identical) +
/// memidx LEB when the 0x40 flag is set.
pub fn simd_memarg_size(code: &[u8], operand_start: usize) -> usize {
    let mut ip = operand_start;
    let align = read_leb_u32(code, &mut ip);
    if align & 0x80 == 0 {
        return 0;
    }
    let _offset = read_leb_u64(code, &mut ip);
    if align & 0x40 != 0 {
        let _memidx = read_leb_u32(code, &mut ip);
    }
    ip.saturating_sub(operand_start)
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
                _ => None }
        }

        pub(super) fn operand_format(sub: u16) -> super::OperandFormat {
            match sub {
                $( $sub => super::OperandFormat::$fmt, )*
                _ => super::OperandFormat::None }
        }

        /// Reverse of `name`: the WASM mnemonic → sub-opcode. Generated from the
        /// same table so there is exactly one opcode list to maintain.
        #[allow(dead_code)]
        pub(super) fn from_name(wasm_name: &str) -> Option<u16> {
            match wasm_name {
                $( $wasm_name => Some($sub), )*
                _ => None }
        }
    };
}

// Re-export the macro for category files
pub(crate) use opcode_category;
