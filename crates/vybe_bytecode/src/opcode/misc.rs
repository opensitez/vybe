//! Misc proposal opcodes (prefix 0xFC).

use super::Op;
use super::opcode_category;

impl Op {
    pub const MEMORY_INIT: Op = Op::new(0xFC, 0x08);
}

opcode_category! {
    [0x08] memory_init => None, "memory.init";
}
