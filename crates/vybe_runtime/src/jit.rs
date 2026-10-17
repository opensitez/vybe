//! Baseline executable tier for a bounded standard-Wasm integer subset.
//! Unsupported code is rejected before executable memory is allocated.
use crate::native_gc::{CompositeType, ValueType};
use crate::{Chunk, Op};
use std::collections::HashMap;

#[derive(Clone, Copy, Debug, Default)]
pub struct Counters {
    pub compile_attempts: u64,
    pub admitted: u64,
    pub native_entries: u64,
    pub rejected: u64,
    pub fuel_fallbacks: u64,
    pub code_bytes: usize,
}

#[derive(Default)]
pub struct State {
    cache: HashMap<usize, Cached>,
    scratch: Vec<i32>,
    pub counters: Counters,
}

struct Cached {
    body: Vec<u8>,
    signature: crate::native_gc::TypeIdentity,
    locals: u16,
    code: Option<Compiled>,
}

impl State {
    pub fn invoke(&mut self, index: usize, chunk: &Chunk, args: &[crate::Value]) -> Option<i32> {
        if !cfg!(all(target_os = "macos", target_arch = "aarch64")) {
            return None;
        }
        let signature = chunk.native_func_type.as_ref()?;
        let CompositeType::Func { params, results } = &signature.definition().composite else {
            return None;
        };
        if results != &[ValueType::Scalar(0x7f)]
            || params.iter().any(|p| p != &ValueType::Scalar(0x7f))
            || params.len() != args.len()
            || params.len() != usize::from(chunk.arity)
            || chunk.result_arity != 1
            || chunk.param_count != chunk.arity
            || chunk.is_async
            || chunk.is_generator
            || chunk.capture_count != 0
            || chunk.takes_receiver
            || args.iter().any(|a| !matches!(a, crate::Value::I32(_)))
        {
            return None;
        }
        let fresh = self.cache.get(&index).is_none_or(|c| {
            c.body != chunk.code || c.signature != *signature || c.locals != chunk.local_count
        });
        if fresh {
            if self.cache.len() >= 128 && !self.cache.contains_key(&index) {
                if let Some(key) = self.cache.keys().next().copied() {
                    if let Some(code) = self.cache.remove(&key).and_then(|c| c.code) {
                        self.counters.code_bytes -= code.memory.len;
                    }
                }
            }
            self.counters.compile_attempts += 1;
            let code = compile(chunk);
            if let Some(code) = &code {
                self.counters.admitted += 1;
                self.counters.code_bytes += code.memory.len;
            } else {
                self.counters.rejected += 1;
            }
            let previous = self.cache.insert(
                index,
                Cached {
                    body: chunk.code.clone(),
                    signature: signature.clone(),
                    locals: chunk.local_count,
                    code,
                },
            );
            if let Some(previous) = previous.and_then(|p| p.code) {
                self.counters.code_bytes -= previous.memory.len;
            }
        }
        let code = self.cache.get(&index)?.code.as_ref()?;
        self.scratch.resize(code.slots, 0);
        self.scratch.fill(0);
        for (slot, value) in self.scratch.iter_mut().zip(args) {
            let crate::Value::I32(n) = value else {
                unreachable!()
            };
            *slot = *n;
        }
        self.counters.native_entries += 1;
        // The admitted program uses only in-bounds i32 scratch slots and
        // caller-saved registers. Its mapping is immutable RX for this call.
        let result = unsafe { code.memory.call(self.scratch.as_mut_ptr()) };
        if self.scratch[code.slots - 1] != 0 {
            self.counters.fuel_fallbacks += 1;
            // The caller restarts the original untouched VM frame once.
            // Do not retry this long-running body until its exact version
            // changes; otherwise every call would repay the native budget.
            if let Some(code) = self.cache.get_mut(&index).and_then(|c| c.code.take()) {
                self.counters.code_bytes -= code.memory.len;
            }
            None
        } else {
            Some(result)
        }
    }
}

struct Compiled {
    memory: Executable,
    slots: usize,
}

struct Label {
    body: usize,
    end: usize,
    height: usize,
    is_loop: bool,
}

fn u16_at(code: &[u8], p: usize) -> Option<u16> {
    Some(u16::from_be_bytes(code.get(p..p + 2)?.try_into().ok()?))
}
fn uleb(code: &[u8], p: &mut usize) -> Option<usize> {
    let mut n = 0u32;
    for shift in (0..35).step_by(7) {
        let b = *code.get(*p)?;
        *p += 1;
        if shift == 28 && b & 0xf0 != 0 {
            return None;
        }
        n |= u32::from(b & 127) << shift;
        if b & 128 == 0 {
            return Some(n as usize);
        }
    }
    None
}
fn sleb32(code: &[u8], p: &mut usize) -> Option<i32> {
    let mut n = 0i64;
    for shift in (0..35).step_by(7) {
        let b = *code.get(*p)?;
        *p += 1;
        n |= i64::from(b & 127) << shift;
        if b & 128 == 0 {
            if b & 64 != 0 {
                n |= (!0i64) << (shift + 7);
            }
            return n.try_into().ok();
        }
    }
    None
}

fn compile(chunk: &Chunk) -> Option<Compiled> {
    if !cfg!(all(target_os = "macos", target_arch = "aarch64")) {
        return None;
    }
    let code = &chunk.code;
    if code.len() > 65536 || chunk.local_count > 256 {
        return None;
    }
    let locals = usize::from(chunk.local_count).max(usize::from(chunk.arity));
    // Bound each native activation. A pure admitted function can safely
    // restart in the interpreter when this budget expires: no guest memory,
    // host calls, reference writes or other externally visible effects occur.
    let mut words = vec![
        0x5280000b | ((100000u32 & 65535) << 5),
        0x72a0000b | ((100000u32 >> 16) << 5),
    ];
    let mut locations = HashMap::new();
    let mut branches = Vec::new();
    let mut bailouts = Vec::new();
    let mut labels: Vec<Label> = Vec::new();
    let mut all_labels = Vec::new();
    let mut open = Vec::new();
    let mut p = 0;
    // Validate the bounded opcode language and find exact structured targets.
    while p < code.len() {
        let at = p;
        code.get(p..p + 4)?;
        let op = Op::from_code_at(code, p);
        p += 4;
        match op {
            Op::BLOCK | Op::LOOP => {
                if code.get(p..p + 2)? != [0, 0] {
                    return None;
                }
                p += 2;
                let id = all_labels.len();
                all_labels.push((at, p, 0, op == Op::LOOP));
                open.push(id);
            }
            Op::END => {
                let id = open.pop()?;
                all_labels[id].2 = p;
            }
            Op::LOCAL_GET | Op::LOCAL_SET | Op::LOCAL_TEE => {
                u16_at(code, p)?;
                p += 2;
            }
            Op::I32_CONST => {
                sleb32(code, &mut p)?;
            }
            Op::BR | Op::BR_IF => {
                uleb(code, &mut p)?;
            }
            Op::I32_ADD
            | Op::I32_SUB
            | Op::I32_MUL
            | Op::I32_AND
            | Op::I32_OR
            | Op::I32_XOR
            | Op::I32_SHL
            | Op::I32_SHR_S
            | Op::I32_SHR_U
            | Op::I32_EQZ
            | Op::I32_EQ
            | Op::I32_NE
            | Op::I32_LT_S
            | Op::I32_LT_U
            | Op::I32_GT_S
            | Op::I32_GT_U
            | Op::I32_LE_S
            | Op::I32_LE_U
            | Op::I32_GE_S
            | Op::I32_GE_U
            | Op::DROP
            | Op::NOP
            | Op::RETURN => {}
            _ => return None,
        }
    }
    if !open.is_empty() {
        return None;
    }
    p = 0;
    let mut height = 0usize;
    let mut maximum = locals;
    let mut returned = false;
    while p < code.len() {
        let at = p;
        locations.insert(at, words.len());
        let op = Op::from_code_at(code, p);
        p += 4;
        let load = |out: &mut Vec<u32>, r: u32, slot: usize| -> Option<()> {
            if slot >= 4096 {
                return None;
            }
            out.push(0xb9400000 | ((slot as u32) << 10) | r);
            Some(())
        };
        let store = |out: &mut Vec<u32>, r: u32, slot: usize| -> Option<()> {
            if slot >= 4096 {
                return None;
            }
            out.push(0xb9000000 | ((slot as u32) << 10) | r);
            Some(())
        };
        match op {
            Op::BLOCK | Op::LOOP => {
                p += 2;
                let &(_, body, end, is_loop) = all_labels.iter().find(|l| l.0 == at)?;
                labels.push(Label {
                    body,
                    end,
                    height,
                    is_loop,
                });
            }
            Op::END => {
                if labels.pop()?.height != height {
                    return None;
                }
            }
            Op::I32_CONST => {
                let n = sleb32(code, &mut p)? as u32;
                words.push(0x52800009 | ((n & 65535) << 5));
                words.push(0x72a00009 | ((n >> 16) << 5));
                store(&mut words, 9, locals + height)?;
                height += 1;
            }
            Op::LOCAL_GET => {
                let i = usize::from(u16_at(code, p)?);
                p += 2;
                if i >= locals {
                    return None;
                }
                load(&mut words, 9, i)?;
                store(&mut words, 9, locals + height)?;
                height += 1;
            }
            Op::LOCAL_SET | Op::LOCAL_TEE => {
                let i = usize::from(u16_at(code, p)?);
                p += 2;
                if i >= locals || height == 0 {
                    return None;
                }
                load(&mut words, 9, locals + height - 1)?;
                store(&mut words, 9, i)?;
                if op == Op::LOCAL_SET {
                    height -= 1;
                }
            }
            Op::I32_EQZ => {
                if height == 0 {
                    return None;
                }
                load(&mut words, 9, locals + height - 1)?;
                words.extend_from_slice(&[0x7100013f, 0x1a9f17e9]);
                store(&mut words, 9, locals + height - 1)?;
            }
            Op::DROP => {
                height = height.checked_sub(1)?;
            }
            Op::NOP => {}
            Op::BR | Op::BR_IF => {
                let depth = uleb(code, &mut p)?;
                if op == Op::BR_IF {
                    height = height.checked_sub(1)?;
                    load(&mut words, 9, locals + height)?;
                }
                let label = labels.get(labels.len().checked_sub(depth + 1)?)?;
                if height != label.height {
                    return None;
                }
                if op == Op::BR
                    && (code.get(p..p + 4).is_none() || Op::from_code_at(code, p) != Op::END)
                {
                    return None;
                }
                if label.is_loop {
                    words.push(0x7100056b); // subs w11,w11,#1
                    bailouts.push(words.len());
                    words.push(0x54000000); // b.eq bounded activation exit
                }
                branches.push((
                    words.len(),
                    if label.is_loop { label.body } else { label.end },
                    op == Op::BR_IF,
                ));
                words.push(if op == Op::BR_IF {
                    0x35000009
                } else {
                    0x14000000
                });
            }
            Op::RETURN => {
                if !labels.is_empty() || height != 1 || p != code.len() {
                    return None;
                }
                load(&mut words, 9, locals)?;
                words.extend_from_slice(&[0x2a0903e0, 0xd65f03c0]);
                returned = true;
            }
            _ => {
                height = height.checked_sub(1)?;
                if height == 0 {
                    return None;
                }
                load(&mut words, 9, locals + height - 1)?;
                load(&mut words, 10, locals + height)?;
                let arithmetic = match op {
                    Op::I32_ADD => Some(0x0b000000),
                    Op::I32_SUB => Some(0x4b000000),
                    Op::I32_MUL => Some(0x1b007c00),
                    Op::I32_AND => Some(0x0a000000),
                    Op::I32_OR => Some(0x2a000000),
                    Op::I32_XOR => Some(0x4a000000),
                    Op::I32_SHL => Some(0x1ac02000),
                    Op::I32_SHR_U => Some(0x1ac02400),
                    Op::I32_SHR_S => Some(0x1ac02800),
                    _ => None,
                };
                if let Some(base) = arithmetic {
                    words.push(base | (10 << 16) | (9 << 5) | 9);
                } else {
                    let cond = match op {
                        Op::I32_EQ => 0,
                        Op::I32_NE => 1,
                        Op::I32_LT_S => 11,
                        Op::I32_LT_U => 3,
                        Op::I32_GT_S => 12,
                        Op::I32_GT_U => 8,
                        Op::I32_LE_S => 13,
                        Op::I32_LE_U => 9,
                        Op::I32_GE_S => 10,
                        Op::I32_GE_U => 2,
                        _ => return None,
                    };
                    words.push(0x6b0a013f);
                    words.push(0x1a9f07e0 | ((cond ^ 1) << 12) | 9);
                }
                store(&mut words, 9, locals + height - 1)?;
            }
        }
        maximum = maximum.max(locals + height);
    }
    if !returned {
        return None;
    }
    locations.insert(code.len(), words.len());
    if maximum >= 4096 {
        return None;
    }
    let bailout = words.len();
    words.extend_from_slice(&[
        0x52800029,
        0xb9000009 | ((maximum as u32) << 10),
        0x52800000,
        0xd65f03c0,
    ]);
    for site in bailouts {
        let offset = bailout as i64 - site as i64;
        if !(0..(1 << 18)).contains(&offset) {
            return None;
        }
        words[site] |= (offset as u32) << 5;
    }
    for (site, target, conditional) in branches {
        let offset = i64::try_from(*locations.get(&target)?).ok()? - site as i64;
        if conditional {
            if !(-(1 << 18)..(1 << 18)).contains(&offset) {
                return None;
            }
            words[site] |= ((offset as u32) & 0x7ffff) << 5;
        } else {
            if !(-(1 << 25)..(1 << 25)).contains(&offset) {
                return None;
            }
            words[site] |= (offset as u32) & 0x3ffffff;
        }
    }
    let bytes: Vec<_> = words.into_iter().flat_map(u32::to_le_bytes).collect();
    Some(Compiled {
        memory: Executable::new(&bytes)?,
        slots: maximum + 1,
    })
}

struct Executable {
    ptr: *mut std::ffi::c_void,
    len: usize,
}
// Only immutable RX code is shared; scratch belongs to the invoking State.
unsafe impl Send for Executable {}
unsafe impl Sync for Executable {}

#[cfg(all(target_os = "macos", target_arch = "aarch64"))]
unsafe extern "C" {
    fn mmap(
        addr: *mut std::ffi::c_void,
        len: usize,
        prot: i32,
        flags: i32,
        fd: i32,
        offset: isize,
    ) -> *mut std::ffi::c_void;
    fn mprotect(addr: *mut std::ffi::c_void, len: usize, prot: i32) -> i32;
    fn munmap(addr: *mut std::ffi::c_void, len: usize) -> i32;
}
#[cfg(all(target_os = "macos", target_arch = "aarch64"))]
unsafe extern "C" {
    fn sys_icache_invalidate(start: *mut std::ffi::c_void, len: usize);
}

impl Executable {
    fn new(bytes: &[u8]) -> Option<Self> {
        #[cfg(all(target_os = "macos", target_arch = "aarch64"))]
        unsafe {
            let ptr = mmap(std::ptr::null_mut(), bytes.len(), 3, 0x1002, -1, 0);
            if ptr as isize == -1 {
                return None;
            }
            std::ptr::copy_nonoverlapping(bytes.as_ptr(), ptr.cast(), bytes.len());
            sys_icache_invalidate(ptr, bytes.len());
            if mprotect(ptr, bytes.len(), 5) != 0 {
                munmap(ptr, bytes.len());
                return None;
            }
            return Some(Self {
                ptr,
                len: bytes.len(),
            });
        }
        #[cfg(not(all(target_os = "macos", target_arch = "aarch64")))]
        {
            let _ = bytes;
            None
        }
    }
    unsafe fn call(&self, scratch: *mut i32) -> i32 {
        let entry: unsafe extern "C" fn(*mut i32) -> i32 = unsafe { std::mem::transmute(self.ptr) };
        unsafe { entry(scratch) }
    }
}
impl Drop for Executable {
    fn drop(&mut self) {
        #[cfg(all(target_os = "macos", target_arch = "aarch64"))]
        unsafe {
            munmap(self.ptr, self.len);
        }
    }
}
