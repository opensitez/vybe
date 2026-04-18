//! WASM-compliant opcode definitions.
//!
//! `Op` is a `u16` newtype: `(prefix << 8) | sub_opcode`.
//! Every opcode is 2 bytes on the wire. No special-casing.
//!
//! Prefixes match the WASM spec:
//! - `0x00`: Core MVP
//! - `0xFB`: GC proposal
//! - `0xFC`: Misc proposal
//! - `0xFD`: SIMD proposal
//! - `0xFE`: Threads proposal
//! - `0xFF`: VM-internal (not WASM — lowered in .wasm output)
//!
//! Opcodes are defined in category files (core.rs, gc.rs, etc.) as `pub const` values.
//! Adding an opcode = one line in one file.

mod core_ops;
mod gc;
mod misc;
mod simd;
mod threads;
mod vm_internal;

/// A bytecode opcode. Encoded as `(prefix << 8) | sub_opcode`.
/// All opcodes are uniformly 2 bytes on the wire.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Op(pub u16);

impl Op {
    /// Create from prefix + sub-opcode.
    #[inline]
    pub const fn new(prefix: u8, sub: u8) -> Self {
        Op(((prefix as u16) << 8) | sub as u16)
    }

    /// Prefix byte (0x00=core, 0xFB=GC, 0xFC=misc, 0xFD=SIMD, 0xFE=threads, 0xFF=VM-internal).
    #[inline]
    pub const fn prefix(self) -> u8 { (self.0 >> 8) as u8 }

    /// Sub-opcode byte within the prefix group.
    #[inline]
    pub const fn sub(self) -> u8 { (self.0 & 0xFF) as u8 }

    /// Encode to 2 bytes: [prefix, sub_opcode]. Uniform encoding.
    #[inline]
    pub const fn encode(self) -> (u8, u8) { (self.prefix(), self.sub()) }

    /// All opcodes are uniformly 2 bytes.
    #[inline]
    pub const fn encoded_len(self) -> usize { 2 }

    /// True if this is a VM-internal opcode (0xFF prefix), not standard WASM.
    #[inline]
    pub const fn is_vm_internal(self) -> bool { self.prefix() == 0xFF }

    /// Decode 2 bytes into a validated opcode.
    /// Returns None if the prefix/sub combination is not a defined opcode.
    pub fn decode(prefix: u8, sub: u8) -> Option<Op> {
        let op = Op::new(prefix, sub);
        // Check if this is a valid opcode by looking up metadata
        if op.wasm_name_opt().is_some() {
            Some(op)
        } else {
            None
        }
    }

    /// Operand format for this opcode.
    pub fn operand_format(self) -> OperandFormat {
        match self.prefix() {
            0x00 => core_ops::operand_format(self.sub()),
            0xFB => gc::operand_format(self.sub()),
            0xFC => misc::operand_format(self.sub()),
            0xFD => simd::operand_format(self.sub()),
            0xFE => threads::operand_format(self.sub()),
            0xFF => vm_internal::operand_format(self.sub()),
            _ => OperandFormat::None,
        }
    }

    /// WASM disassembly name, or None if not a valid opcode.
    pub fn wasm_name_opt(self) -> Option<&'static str> {
        match self.prefix() {
            0x00 => core_ops::name(self.sub()),
            0xFB => gc::name(self.sub()),
            0xFC => misc::name(self.sub()),
            0xFD => simd::name(self.sub()),
            0xFE => threads::name(self.sub()),
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
    /// Variable: u16 func_idx + u8 upvalue_count + descriptors.
    Closure,
    /// Variable: u8 count + u8 default + count × u8 labels.
    BrTable,
    /// Variable: u8 handler_count + handlers.
    TryTable,
    /// 16 bytes immediate (v128.const).
    V128Const,
    /// 16 bytes lane indices (i8x16.shuffle).
    Shuffle,
}

impl OperandFormat {
    /// Fixed operand size in bytes, or 0 for variable-length formats.
    pub const fn fixed_size(self) -> usize {
        match self {
            Self::None => 0,
            Self::U8 => 1,
            Self::U16 | Self::I16 => 2,
            Self::U16_U8 => 3,
            Self::U16_U16 | Self::U16_I16 => 4,
            Self::V128Const | Self::Shuffle => 16,
            Self::Closure | Self::BrTable | Self::TryTable => 0,
        }
    }
}

/// Helper macro for category files. Generates name() and operand_format() functions
/// from a table of [sub] name => format entries.
macro_rules! opcode_category {
    ( $( [$sub:literal] $name:ident => $fmt:ident, $wasm_name:literal; )* ) => {
        pub(super) fn name(sub: u8) -> Option<&'static str> {
            match sub {
                $( $sub => Some($wasm_name), )*
                _ => None,
            }
        }

        pub(super) fn operand_format(sub: u8) -> super::OperandFormat {
            match sub {
                $( $sub => super::OperandFormat::$fmt, )*
                _ => super::OperandFormat::None,
            }
        }
    };
}

// Re-export the macro for category files
pub(crate) use opcode_category;
