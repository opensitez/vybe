//! Runtime semantics for the WASM threads/atomics proposal.
//!
//! The VM keeps opcodes fixed-width internally, but opcode identity remains
//! the WASM `(0xFE, subopcode)` pair. Any instruction immediates, such as
//! memarg, follow those two opcode bytes in the chunk code.

use crate::error::VMError;
use crate::opcode::{Op, read_leb_u32, read_leb_u64};
use crate::value::Value;
use crate::vm::VM;

#[derive(Clone, Copy)]
enum AtomicWidth {
    W8,
    W16,
    W32,
    W64,
}

impl AtomicWidth {
    fn bytes(self) -> usize {
        match self {
            Self::W8 => 1,
            Self::W16 => 2,
            Self::W32 => 4,
            Self::W64 => 8,
        }
    }
}

#[derive(Clone, Copy)]
enum AtomicRmw {
    Add,
    Sub,
    And,
    Or,
    Xor,
    Xchg,
}

impl VM {
    pub(crate) fn execute_threads_op(&mut self, op: Op) -> Result<bool, VMError> {
        match op {
            _ if op == Op::ATOMIC_FENCE => {
                self.read_atomic_fence_immediate();
                Ok(true)
            }
            _ if op == Op::MEMORY_ATOMIC_NOTIFY => {
                let count = self.pop().as_i32();
                let (_, addr) = self.pop_atomic_addr(AtomicWidth::W32)?;
                self.push(Value::I32(self.memory.notify(addr, count)))?;
                Ok(true)
            }
            _ if op == Op::MEMORY_ATOMIC_WAIT32 => {
                let timeout = self.pop().as_i64();
                let expected = self.pop().as_i32();
                let (_, addr) = self.pop_atomic_addr(AtomicWidth::W32)?;
                self.push(Value::I32(self.memory.wait32(addr, expected, timeout)))?;
                Ok(true)
            }
            _ if op == Op::MEMORY_ATOMIC_WAIT64 => {
                let timeout = self.pop().as_i64();
                let expected = self.pop().as_i64();
                let (_, addr) = self.pop_atomic_addr(AtomicWidth::W64)?;
                self.push(Value::I32(self.memory.wait64(addr, expected, timeout)))?;
                Ok(true)
            }
            _ if op == Op::I32_ATOMIC_LOAD => {
                let (memidx, addr) = self.pop_atomic_addr(AtomicWidth::W32)?;
                let raw = self.atomic_load(memidx, addr, AtomicWidth::W32)?;
                self.push(Value::I32(raw as u32 as i32))?;
                Ok(true)
            }
            _ if op == Op::I64_ATOMIC_LOAD => {
                let (memidx, addr) = self.pop_atomic_addr(AtomicWidth::W64)?;
                let raw = self.atomic_load(memidx, addr, AtomicWidth::W64)?;
                self.push(Value::I64(raw as i64))?;
                Ok(true)
            }
            _ if op == Op::I32_ATOMIC_LOAD8_U => {
                let (memidx, addr) = self.pop_atomic_addr(AtomicWidth::W8)?;
                self.push(Value::I32(
                    self.atomic_load(memidx, addr, AtomicWidth::W8)? as i32
                ))?;
                Ok(true)
            }
            _ if op == Op::I32_ATOMIC_LOAD16_U => {
                let (memidx, addr) = self.pop_atomic_addr(AtomicWidth::W16)?;
                self.push(Value::I32(
                    self.atomic_load(memidx, addr, AtomicWidth::W16)? as i32,
                ))?;
                Ok(true)
            }
            _ if op == Op::I64_ATOMIC_LOAD8_U => {
                let (memidx, addr) = self.pop_atomic_addr(AtomicWidth::W8)?;
                self.push(Value::I64(
                    self.atomic_load(memidx, addr, AtomicWidth::W8)? as i64
                ))?;
                Ok(true)
            }
            _ if op == Op::I64_ATOMIC_LOAD16_U => {
                let (memidx, addr) = self.pop_atomic_addr(AtomicWidth::W16)?;
                self.push(Value::I64(
                    self.atomic_load(memidx, addr, AtomicWidth::W16)? as i64,
                ))?;
                Ok(true)
            }
            _ if op == Op::I64_ATOMIC_LOAD32_U => {
                let (memidx, addr) = self.pop_atomic_addr(AtomicWidth::W32)?;
                self.push(Value::I64(
                    self.atomic_load(memidx, addr, AtomicWidth::W32)? as i64,
                ))?;
                Ok(true)
            }
            _ if op == Op::I32_ATOMIC_STORE => self.atomic_store_i32(AtomicWidth::W32),
            _ if op == Op::I32_ATOMIC_STORE8 => self.atomic_store_i32(AtomicWidth::W8),
            _ if op == Op::I32_ATOMIC_STORE16 => self.atomic_store_i32(AtomicWidth::W16),
            _ if op == Op::I64_ATOMIC_STORE => self.atomic_store_i64(AtomicWidth::W64),
            _ if op == Op::I64_ATOMIC_STORE8 => self.atomic_store_i64(AtomicWidth::W8),
            _ if op == Op::I64_ATOMIC_STORE16 => self.atomic_store_i64(AtomicWidth::W16),
            _ if op == Op::I64_ATOMIC_STORE32 => self.atomic_store_i64(AtomicWidth::W32),
            _ if op == Op::I32_ATOMIC_RMW_ADD => {
                self.atomic_rmw_i32(AtomicWidth::W32, AtomicRmw::Add)
            }
            _ if op == Op::I32_ATOMIC_RMW_SUB => {
                self.atomic_rmw_i32(AtomicWidth::W32, AtomicRmw::Sub)
            }
            _ if op == Op::I32_ATOMIC_RMW_AND => {
                self.atomic_rmw_i32(AtomicWidth::W32, AtomicRmw::And)
            }
            _ if op == Op::I32_ATOMIC_RMW_OR => {
                self.atomic_rmw_i32(AtomicWidth::W32, AtomicRmw::Or)
            }
            _ if op == Op::I32_ATOMIC_RMW_XOR => {
                self.atomic_rmw_i32(AtomicWidth::W32, AtomicRmw::Xor)
            }
            _ if op == Op::I32_ATOMIC_RMW_XCHG => {
                self.atomic_rmw_i32(AtomicWidth::W32, AtomicRmw::Xchg)
            }
            _ if op == Op::I32_ATOMIC_RMW8_ADD_U => {
                self.atomic_rmw_i32(AtomicWidth::W8, AtomicRmw::Add)
            }
            _ if op == Op::I32_ATOMIC_RMW16_ADD_U => {
                self.atomic_rmw_i32(AtomicWidth::W16, AtomicRmw::Add)
            }
            _ if op == Op::I32_ATOMIC_RMW8_SUB_U => {
                self.atomic_rmw_i32(AtomicWidth::W8, AtomicRmw::Sub)
            }
            _ if op == Op::I32_ATOMIC_RMW16_SUB_U => {
                self.atomic_rmw_i32(AtomicWidth::W16, AtomicRmw::Sub)
            }
            _ if op == Op::I32_ATOMIC_RMW8_AND_U => {
                self.atomic_rmw_i32(AtomicWidth::W8, AtomicRmw::And)
            }
            _ if op == Op::I32_ATOMIC_RMW16_AND_U => {
                self.atomic_rmw_i32(AtomicWidth::W16, AtomicRmw::And)
            }
            _ if op == Op::I32_ATOMIC_RMW8_OR_U => {
                self.atomic_rmw_i32(AtomicWidth::W8, AtomicRmw::Or)
            }
            _ if op == Op::I32_ATOMIC_RMW16_OR_U => {
                self.atomic_rmw_i32(AtomicWidth::W16, AtomicRmw::Or)
            }
            _ if op == Op::I32_ATOMIC_RMW8_XOR_U => {
                self.atomic_rmw_i32(AtomicWidth::W8, AtomicRmw::Xor)
            }
            _ if op == Op::I32_ATOMIC_RMW16_XOR_U => {
                self.atomic_rmw_i32(AtomicWidth::W16, AtomicRmw::Xor)
            }
            _ if op == Op::I32_ATOMIC_RMW8_XCHG_U => {
                self.atomic_rmw_i32(AtomicWidth::W8, AtomicRmw::Xchg)
            }
            _ if op == Op::I32_ATOMIC_RMW16_XCHG_U => {
                self.atomic_rmw_i32(AtomicWidth::W16, AtomicRmw::Xchg)
            }
            _ if op == Op::I64_ATOMIC_RMW_ADD => {
                self.atomic_rmw_i64(AtomicWidth::W64, AtomicRmw::Add)
            }
            _ if op == Op::I64_ATOMIC_RMW_SUB => {
                self.atomic_rmw_i64(AtomicWidth::W64, AtomicRmw::Sub)
            }
            _ if op == Op::I64_ATOMIC_RMW_AND => {
                self.atomic_rmw_i64(AtomicWidth::W64, AtomicRmw::And)
            }
            _ if op == Op::I64_ATOMIC_RMW_OR => {
                self.atomic_rmw_i64(AtomicWidth::W64, AtomicRmw::Or)
            }
            _ if op == Op::I64_ATOMIC_RMW_XOR => {
                self.atomic_rmw_i64(AtomicWidth::W64, AtomicRmw::Xor)
            }
            _ if op == Op::I64_ATOMIC_RMW_XCHG => {
                self.atomic_rmw_i64(AtomicWidth::W64, AtomicRmw::Xchg)
            }
            _ if op == Op::I64_ATOMIC_RMW8_ADD_U => {
                self.atomic_rmw_i64(AtomicWidth::W8, AtomicRmw::Add)
            }
            _ if op == Op::I64_ATOMIC_RMW16_ADD_U => {
                self.atomic_rmw_i64(AtomicWidth::W16, AtomicRmw::Add)
            }
            _ if op == Op::I64_ATOMIC_RMW32_ADD_U => {
                self.atomic_rmw_i64(AtomicWidth::W32, AtomicRmw::Add)
            }
            _ if op == Op::I64_ATOMIC_RMW8_SUB_U => {
                self.atomic_rmw_i64(AtomicWidth::W8, AtomicRmw::Sub)
            }
            _ if op == Op::I64_ATOMIC_RMW16_SUB_U => {
                self.atomic_rmw_i64(AtomicWidth::W16, AtomicRmw::Sub)
            }
            _ if op == Op::I64_ATOMIC_RMW32_SUB_U => {
                self.atomic_rmw_i64(AtomicWidth::W32, AtomicRmw::Sub)
            }
            _ if op == Op::I64_ATOMIC_RMW8_AND_U => {
                self.atomic_rmw_i64(AtomicWidth::W8, AtomicRmw::And)
            }
            _ if op == Op::I64_ATOMIC_RMW16_AND_U => {
                self.atomic_rmw_i64(AtomicWidth::W16, AtomicRmw::And)
            }
            _ if op == Op::I64_ATOMIC_RMW32_AND_U => {
                self.atomic_rmw_i64(AtomicWidth::W32, AtomicRmw::And)
            }
            _ if op == Op::I64_ATOMIC_RMW8_OR_U => {
                self.atomic_rmw_i64(AtomicWidth::W8, AtomicRmw::Or)
            }
            _ if op == Op::I64_ATOMIC_RMW16_OR_U => {
                self.atomic_rmw_i64(AtomicWidth::W16, AtomicRmw::Or)
            }
            _ if op == Op::I64_ATOMIC_RMW32_OR_U => {
                self.atomic_rmw_i64(AtomicWidth::W32, AtomicRmw::Or)
            }
            _ if op == Op::I64_ATOMIC_RMW8_XOR_U => {
                self.atomic_rmw_i64(AtomicWidth::W8, AtomicRmw::Xor)
            }
            _ if op == Op::I64_ATOMIC_RMW16_XOR_U => {
                self.atomic_rmw_i64(AtomicWidth::W16, AtomicRmw::Xor)
            }
            _ if op == Op::I64_ATOMIC_RMW32_XOR_U => {
                self.atomic_rmw_i64(AtomicWidth::W32, AtomicRmw::Xor)
            }
            _ if op == Op::I64_ATOMIC_RMW8_XCHG_U => {
                self.atomic_rmw_i64(AtomicWidth::W8, AtomicRmw::Xchg)
            }
            _ if op == Op::I64_ATOMIC_RMW16_XCHG_U => {
                self.atomic_rmw_i64(AtomicWidth::W16, AtomicRmw::Xchg)
            }
            _ if op == Op::I64_ATOMIC_RMW32_XCHG_U => {
                self.atomic_rmw_i64(AtomicWidth::W32, AtomicRmw::Xchg)
            }
            _ if op == Op::I32_ATOMIC_RMW_CMPXCHG => self.atomic_cmpxchg_i32(AtomicWidth::W32),
            _ if op == Op::I32_ATOMIC_RMW8_CMPXCHG_U => self.atomic_cmpxchg_i32(AtomicWidth::W8),
            _ if op == Op::I32_ATOMIC_RMW16_CMPXCHG_U => self.atomic_cmpxchg_i32(AtomicWidth::W16),
            _ if op == Op::I64_ATOMIC_RMW_CMPXCHG => self.atomic_cmpxchg_i64(AtomicWidth::W64),
            _ if op == Op::I64_ATOMIC_RMW8_CMPXCHG_U => self.atomic_cmpxchg_i64(AtomicWidth::W8),
            _ if op == Op::I64_ATOMIC_RMW16_CMPXCHG_U => self.atomic_cmpxchg_i64(AtomicWidth::W16),
            _ if op == Op::I64_ATOMIC_RMW32_CMPXCHG_U => self.atomic_cmpxchg_i64(AtomicWidth::W32),
            _ => Ok(false),
        }
    }

    fn read_atomic_fence_immediate(&mut self) {
        let _ = self.read_byte();
    }

    fn pop_atomic_addr(&mut self, width: AtomicWidth) -> Result<(usize, usize), VMError> {
        let (offset, memidx, memory64) = self.read_atomic_memarg();
        let addr = if memory64 {
            let base = self.pop().as_i64();
            if base < 0 {
                return Err(VMError::new("trap: atomic memory64 negative address"));
            }
            let addr = (base as u64)
                .checked_add(offset)
                .ok_or_else(|| VMError::new("trap: atomic memory64 address overflow"))?;
            usize::try_from(addr)
                .map_err(|_| VMError::new("trap: atomic memory64 address out of range"))?
        } else {
            let base = self.pop().as_i32() as u32 as usize;
            base.checked_add(offset as usize)
                .ok_or_else(|| VMError::new("trap: atomic address overflow"))?
        };
        self.check_atomic_access(memidx, addr, width)?;
        Ok((memidx, addr))
    }

    fn read_atomic_memarg(&mut self) -> (u64, usize, bool) {
        let chunk_idx = self.frame().chunk_index;
        let code = &self.chunks[chunk_idx].code;
        let mut ip = self.frame().ip;
        let align = read_leb_u32(code, &mut ip);
        let memory64 = align & 0x80 != 0;
        let offset = if memory64 {
            read_leb_u64(code, &mut ip)
        } else {
            read_leb_u32(code, &mut ip) as u64
        };
        let memidx = if align & 0x40 != 0 {
            read_leb_u32(code, &mut ip) as usize
        } else {
            0
        };
        self.frame_mut().ip = ip;
        (offset, memidx, memory64)
    }

    fn check_atomic_access(
        &self,
        memidx: usize,
        addr: usize,
        width: AtomicWidth,
    ) -> Result<(), VMError> {
        let size = width.bytes();
        if addr % size != 0 {
            return Err(VMError::new(format!(
                "trap: atomic unaligned access: addr={} size={}",
                addr, size
            )));
        }
        let limit = self.mem_len(memidx);
        if addr.saturating_add(size) > limit {
            return Err(VMError::new(format!(
                "trap: atomic memory access out of bounds: addr={} size={} limit={}",
                addr, size, limit
            )));
        }
        Ok(())
    }

    fn atomic_load(&self, memidx: usize, addr: usize, width: AtomicWidth) -> Result<u64, VMError> {
        let bytes = self.read_memory_bytes(memidx, addr, width.bytes())?;
        Ok(match width {
            AtomicWidth::W8 => bytes[0] as u64,
            AtomicWidth::W16 => u16::from_le_bytes(bytes.try_into().unwrap()) as u64,
            AtomicWidth::W32 => u32::from_le_bytes(bytes.try_into().unwrap()) as u64,
            AtomicWidth::W64 => u64::from_le_bytes(bytes.try_into().unwrap()),
        })
    }

    fn atomic_store_raw(
        &mut self,
        memidx: usize,
        addr: usize,
        width: AtomicWidth,
        value: u64,
    ) -> Result<(), VMError> {
        match width {
            AtomicWidth::W8 => self.write_memory_bytes(memidx, addr, &[value as u8]),
            AtomicWidth::W16 => {
                self.write_memory_bytes(memidx, addr, &(value as u16).to_le_bytes())
            }
            AtomicWidth::W32 => {
                self.write_memory_bytes(memidx, addr, &(value as u32).to_le_bytes())
            }
            AtomicWidth::W64 => self.write_memory_bytes(memidx, addr, &value.to_le_bytes()),
        }
    }

    fn atomic_store_i32(&mut self, width: AtomicWidth) -> Result<bool, VMError> {
        let value = self.pop().as_i32() as u32 as u64;
        let (memidx, addr) = self.pop_atomic_addr(width)?;
        self.atomic_store_raw(memidx, addr, width, value)?;
        Ok(true)
    }

    fn atomic_store_i64(&mut self, width: AtomicWidth) -> Result<bool, VMError> {
        let value = self.pop().as_i64() as u64;
        let (memidx, addr) = self.pop_atomic_addr(width)?;
        self.atomic_store_raw(memidx, addr, width, value)?;
        Ok(true)
    }

    fn atomic_rmw_i32(&mut self, width: AtomicWidth, op: AtomicRmw) -> Result<bool, VMError> {
        let value = self.pop().as_i32() as u32 as u64;
        let (memidx, addr) = self.pop_atomic_addr(width)?;
        let old = self.atomic_load(memidx, addr, width)?;
        let new = apply_rmw(old, value, width, op);
        self.atomic_store_raw(memidx, addr, width, new)?;
        self.push(Value::I32((old & width_mask(width)) as u32 as i32))?;
        Ok(true)
    }

    fn atomic_rmw_i64(&mut self, width: AtomicWidth, op: AtomicRmw) -> Result<bool, VMError> {
        let value = self.pop().as_i64() as u64;
        let (memidx, addr) = self.pop_atomic_addr(width)?;
        let old = self.atomic_load(memidx, addr, width)?;
        let new = apply_rmw(old, value, width, op);
        self.atomic_store_raw(memidx, addr, width, new)?;
        self.push(Value::I64((old & width_mask(width)) as i64))?;
        Ok(true)
    }

    fn atomic_cmpxchg_i32(&mut self, width: AtomicWidth) -> Result<bool, VMError> {
        let replacement = self.pop().as_i32() as u32 as u64;
        let expected = self.pop().as_i32() as u32 as u64;
        let (memidx, addr) = self.pop_atomic_addr(width)?;
        let old = self.atomic_load(memidx, addr, width)?;
        if old & width_mask(width) == expected & width_mask(width) {
            self.atomic_store_raw(memidx, addr, width, replacement)?;
        }
        self.push(Value::I32((old & width_mask(width)) as u32 as i32))?;
        Ok(true)
    }

    fn atomic_cmpxchg_i64(&mut self, width: AtomicWidth) -> Result<bool, VMError> {
        let replacement = self.pop().as_i64() as u64;
        let expected = self.pop().as_i64() as u64;
        let (memidx, addr) = self.pop_atomic_addr(width)?;
        let old = self.atomic_load(memidx, addr, width)?;
        if old & width_mask(width) == expected & width_mask(width) {
            self.atomic_store_raw(memidx, addr, width, replacement)?;
        }
        self.push(Value::I64((old & width_mask(width)) as i64))?;
        Ok(true)
    }
}

fn width_mask(width: AtomicWidth) -> u64 {
    match width {
        AtomicWidth::W8 => 0xff,
        AtomicWidth::W16 => 0xffff,
        AtomicWidth::W32 => 0xffff_ffff,
        AtomicWidth::W64 => u64::MAX,
    }
}

fn apply_rmw(old: u64, value: u64, width: AtomicWidth, op: AtomicRmw) -> u64 {
    let mask = width_mask(width);
    let old = old & mask;
    let value = value & mask;
    match op {
        AtomicRmw::Add => old.wrapping_add(value) & mask,
        AtomicRmw::Sub => old.wrapping_sub(value) & mask,
        AtomicRmw::And => old & value,
        AtomicRmw::Or => old | value,
        AtomicRmw::Xor => old ^ value,
        AtomicRmw::Xchg => value,
    }
}
