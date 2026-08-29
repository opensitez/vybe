//! Call Tags proposal opcodes.
//!
//! `proposals/call-tags/proposals/call-tags/Overview.md`. Call tags let a
//! `funcref` handle MORE THAN ONE signature, and — the property this codebase
//! needs — let two functions whose wasm signatures are IDENTICAL still be
//! called under distinguishable conventions:
//!
//! > Notice that a `funcref` can handle multiple call tags, even multiple
//! > canonical call tags. This is because dynamic typing is implemented by the
//! > call*ee* rather than the call*er* […] a `funcref` [can] actually provide
//! > different behavior for different call tags.
//!
//! Under the GC proposal function types are structurally canonicalised, so
//! `(func (param externref))` used for two different purposes is ONE type and
//! the intent is unrecoverable. A custom tag made by `call_tag.new` is a fresh
//! identity over the same signature, which is exactly the distinction a
//! structural type system cannot express.
//!
//! ## Byte assignment
//!
//! ⚠ The proposal is at phase 0: its repository contains `Overview.md` and
//! nothing else — no modified spec document, no binary format, no reference
//! interpreter, no assigned prefix or opcode bytes. **The prefix and sub-opcodes
//! below are THIS PROJECT'S, not WebAssembly's**, and must be renumbered if the
//! proposal is ever assigned bytes. The semantics are the Overview's; only the
//! encoding is ours. Prefix `0xF1` is free here: `0x00` core, `0xF0` canon,
//! `0xFB` GC, `0xFC` misc, `0xFD` SIMD, `0xFE` threads, `0xFF` vm-internal.

use super::Op;
use super::opcode_category;

impl Op {
    /// `call_with_tag $call_tag : [ti* funcref] -> [to*]` — call the `funcref`
    /// on top of the stack under `$call_tag`. The callee decides: if it handles
    /// the tag it runs, otherwise the tag's fall-back handler is called with the
    /// funcref in place of the tag, and a tag with no fall-back traps.
    pub const CALL_WITH_TAG: Op = Op::new(0xF1, 0x00);

    /// `call_indirect_with_tag $table $call_tag : [ti* i32] -> [to*]` —
    /// the Overview defines it as shorthand for
    /// `(call_with_tag $call_tag (table.get $table))`.
    pub const CALL_INDIRECT_WITH_TAG: Op = Op::new(0xF1, 0x01);

    /// `call_return_with_tag $call_tag : [ti* funcref] -> [to*]` — the tail-call
    /// form, for engines supporting `return_call`.
    pub const CALL_RETURN_WITH_TAG: Op = Op::new(0xF1, 0x02);
}

// ⛔⛔ THESE WIDTHS WERE WRONG, AND `operand_format` IS NOT JUST FOR PRINTING.
//
// Declared: U16 / I16 / U16. Actually emitted (and actually READ by
// `dispatch.rs`):
//
//   call_with_tag          emit_op_u16(tag) + u8(argc)              = u16 + u8
//   call_return_with_tag   same                                     = u16 + u8
//   call_indirect_with_tag emit_op_u16(table) + u16(tag) + u8(argc) = u16+u16+u8
//
// So every consumer under-advanced by 1 byte (3 for the indirect form) and then
// decoded the ARGC BYTE as the start of the next opcode. `size_in` is used by
// the wasm writer, `globals.rs`'s global-index rewrite, `link.rs`, `polyfills`,
// the VM's own scanners and both disassemblers — so this desynchronised every
// bytecode walk after a call tag, not merely the `--dump` output. It is why
// `--dump` printed `UNKNOWN(...)` followed by a fabricated `if 0 0` and why
// every offset after a `call_with_tag` was fiction.
//
// The VM's dispatch arm is the authority: it reads u16 (+u16) + byte.
opcode_category! {
    [0x00] call_with_tag => U16_U8, "call_with_tag";
    [0x01] call_indirect_with_tag => U16_U16_U8, "call_indirect_with_tag";
    [0x02] call_return_with_tag => U16_U8, "call_return_with_tag";
}
