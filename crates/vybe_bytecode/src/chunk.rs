use crate::opcode::Op;
use crate::value::Value;

/// A compiled chunk of bytecode — one per function/script.
#[derive(Debug, Clone)]
pub struct Chunk {
    pub code: Vec<u8>,
    pub constants: Vec<Value>,
    pub lines: Vec<u32>,
    pub name: String,
    pub arity: u8,
    pub local_count: u16,
}

impl Chunk {
    pub fn new(name: impl Into<String>) -> Self {
        Chunk {
            code: Vec::new(),
            constants: Vec::new(),
            lines: Vec::new(),
            name: name.into(),
            arity: 0,
            local_count: 0,
        }
    }

    pub fn emit(&mut self, byte: u8, line: u32) {
        self.code.push(byte);
        self.lines.push(line);
    }

    pub fn emit_op(&mut self, op: Op, line: u32) {
        self.emit(op as u8, line);
    }

    pub fn emit_op_u16(&mut self, op: Op, operand: u16, line: u32) {
        self.emit(op as u8, line);
        self.emit((operand >> 8) as u8, line);
        self.emit((operand & 0xff) as u8, line);
    }

    pub fn emit_op_u8(&mut self, op: Op, operand: u8, line: u32) {
        self.emit(op as u8, line);
        self.emit(operand, line);
    }

    pub fn add_constant(&mut self, value: Value) -> u16 {
        self.constants.push(value);
        (self.constants.len() - 1) as u16
    }

    pub fn emit_jump(&mut self, op: Op, line: u32) -> usize {
        self.emit(op as u8, line);
        self.emit(0xff, line);
        self.emit(0xff, line);
        self.code.len() - 2
    }

    pub fn patch_jump(&mut self, offset: usize) {
        let jump = self.code.len() as i32 - (offset as i32 + 2);
        self.code[offset] = (jump >> 8) as u8;
        self.code[offset + 1] = (jump & 0xff) as u8;
    }

    pub fn current_offset(&self) -> usize {
        self.code.len()
    }

    pub fn emit_loop(&mut self, target_offset: usize, line: u32) {
        self.emit(Op::Jump as u8, line);
        let jump = target_offset as i32 - (self.code.len() as i32 + 2);
        self.emit((jump >> 8) as u8, line);
        self.emit((jump & 0xff) as u8, line);
    }

    pub fn read_u16(&self, offset: usize) -> u16 {
        ((self.code[offset] as u16) << 8) | (self.code[offset + 1] as u16)
    }

    pub fn read_i16(&self, offset: usize) -> i16 {
        self.read_u16(offset) as i16
    }
}
