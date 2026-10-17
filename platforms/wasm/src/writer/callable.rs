//! Isolated native-only carrier candidate; original Chunks and VM are untouched.
use crate::encoding::*;
use std::collections::BTreeSet;
use vybe_runtime::{Chunk, Op};
#[derive(Default)]
pub struct Plan {
    pub targets: Vec<u32>,
    pub cell: u32,
    pub captures: u32,
    pub carrier: u32,
    pub import_count: u32,
    pub cache_base: u32,
    pub chunks: Vec<(u8, bool)>,
    pub calls: std::collections::BTreeMap<(u8, u8), u32>,
    pub old_imported_globals: u32,
    pub added_imported_globals: u32,
    pub old_global_len: u32,
    pub global_map: Vec<u32>,
    pub locals: std::collections::HashMap<usize, (u32, Vec<u16>)>,
}
pub fn source(c: &Chunk) -> bool {
    c.func_sig.is_none()
}
pub fn uses_carrier_call(c: &Chunk, offset: usize) -> bool {
    source(c) || !c.call_indirect_sigs.contains_key(&offset)
}

#[cfg(test)]
mod call_representation_tests {
    use super::*;

    #[test]
    fn typed_caller_preserves_declared_core_calls_but_unwraps_generic_helpers() {
        let mut caller = Chunk::new("typed-caller");
        caller.func_sig = Some(("i32".into(), "i32".into()));
        assert!(uses_carrier_call(&caller, 12));
        caller.call_indirect_sigs.insert(12, "i32->i32".into());
        assert!(!uses_carrier_call(&caller, 12));
        assert!(uses_carrier_call(&caller, 24));
    }
}
pub fn captured_slots(c: &Chunk) -> Vec<u16> {
    let mut slots = BTreeSet::new();
    let mut p = 0;
    while p + 4 <= c.code.len() {
        let Some(o) = Op::decode(
            u16::from_be_bytes([c.code[p], c.code[p + 1]]),
            u16::from_be_bytes([c.code[p + 2], c.code[p + 3]]),
        ) else {
            panic!("malformed callable IR");
        };
        if o == Op::REF_FUNC {
            let n = (c.code[p + 6] & 127) as usize;
            for i in 0..n {
                let q = p + 7 + i * 3;
                if c.code[q] != 0 {
                    slots.insert(u16::from_be_bytes([c.code[q + 1], c.code[q + 2]]));
                }
            }
        }
        p += super::code::opcode_size(o, &c.code, p);
    }
    slots.into_iter().collect()
}
pub fn local(b: &mut Vec<u8>, op: u8, i: u32) {
    b.push(op);
    write_leb128_u32(b, i);
}
pub fn gc(b: &mut Vec<u8>, op: u32) {
    b.push(0xfb);
    write_leb128_u32(b, op);
}
pub fn typed(b: &mut Vec<u8>, op: u32, t: u32) {
    gc(b, op);
    if matches!(op, 0x14..=0x17) {
        // ref.test/ref.cast encode a signed s33 heap type, unlike GC typeidx operands.
        write_leb128_i64(b, i64::from(t));
    } else {
        write_leb128_u32(b, t);
    }
}
pub fn field(b: &mut Vec<u8>, op: u32, t: u32, f: u32) {
    typed(b, op, t);
    write_leb128_u32(b, f);
}
pub fn internal(b: &mut Vec<u8>, t: u32) {
    gc(b, 0x1a);
    typed(b, 0x16, t);
}
pub fn external(b: &mut Vec<u8>) {
    gc(b, 0x1b);
}
pub fn null(b: &mut Vec<u8>, ht: i32) {
    b.push(0xd0);
    write_leb128_i32(b, ht);
}
pub fn i32c(b: &mut Vec<u8>, v: i32) {
    b.push(0x41);
    write_leb128_i32(b, v);
}
pub fn carrier(
    b: &mut Vec<u8>,
    p: &Plan,
    index: u32,
    caps: impl FnOnce(&mut Vec<u8>),
    meta: (u8, bool),
) {
    b.push(0xd2);
    write_leb128_u32(b, p.targets[index as usize]);
    caps(b);
    i32c(b, 0);
    null(b, -17);
    i32c(b, meta.0 as i32);
    i32c(b, i32::from(!meta.1));
    typed(b, 0, p.carrier);
}
pub fn cell_slot(c: &Chunk, p: &Plan, slot: u16) -> Option<u32> {
    let (temps, slots) = p
        .locals
        .get(&(c as *const Chunk as usize))
        .expect("same borrowed Chunk epoch");
    slots
        .iter()
        .position(|&x| x == slot)
        .map(|i| u32::from(c.local_count).max(u32::from(c.arity)) + *temps + 1 + i as u32)
}
pub fn temps(c: &Chunk, p: &Plan) -> u32 {
    p.locals[&(c as *const Chunk as usize)].0
}
pub fn hidden_virtual(c: &Chunk, temp_count: u32) -> u32 {
    u32::from(c.local_count).max(u32::from(c.arity)) + temp_count
}
pub fn initialize_cells(b: &mut Vec<u8>, c: &Chunk, p: &Plan, temp_count: u32) {
    if !source(c) {
        return;
    }
    let env = hidden_virtual(c, temp_count);
    for i in 0..c.capture_count {
        // Missing capture descriptor leaves the VM capture-local initialized null.
        local(b, 0x20, env);
        b.push(0xd1);
        b.extend([0x04, 0x6f]);
        null(b, -17);
        b.push(0x05);
        local(b, 0x20, env);
        internal(b, p.captures);
        gc(b, 15);
        i32c(b, i as i32);
        b.push(0x4b);
        b.extend([0x04, 0x6f]);
        local(b, 0x20, env);
        internal(b, p.captures);
        i32c(b, i as i32);
        typed(b, 11, p.captures);
        field(b, 2, p.cell, 0);
        b.push(0x05);
        null(b, -17);
        b.extend([0x0b, 0x0b]);
        local(b, 0x21, u32::from(c.capture_base) + u32::from(i));
    }
    for slot in captured_slots(c) {
        assert!(
            !c.i64_locals.contains(&slot),
            "captured raw i64 requires a declared boxing producer"
        );
        local(b, 0x20, slot as u32);
        typed(b, 0, p.cell);
        external(b);
        local(b, 0x21, cell_slot(c, p, slot).unwrap());
    }
}
/// Remap generated native operands, not IR offsets or u16 logical slots.
/// The hidden capture argument precedes ordinary parameters; all old locals and
/// every native temporary are shifted coherently, including peephole temporaries.
pub fn remap_locals(bytes: &[u8], c: &Chunk, temp_count: u32, p: &Plan) -> Vec<u8> {
    let env = hidden_virtual(c, temp_count);

    let mut r = crate::binary_cursor::Cursor::new(bytes);
    let mut out = Vec::with_capacity(bytes.len());
    while r.remaining() != 0 {
        let start = r.position();
        let op = crate::binary_instructions::next(&mut r).unwrap_or_else(|e| {
            panic!(
                "native callable remap chunk {:?} offset {start}: {e:?}; bytes {:02x?}",
                c.name,
                &bytes[start.saturating_sub(12).min(bytes.len())
                    ..bytes.len().min(start.saturating_add(24))]
            )
        });
        let end = op.end;
        let v = op.index.map(|index| (op.opcode, index));
        if let Some((op, i)) = v {
            let index = if op == 0x23 || op == 0x24 {
                if i & 0x8000_0000 != 0 {
                    i & 0x7fff_ffff
                } else if let Some(&mapped) = p.global_map.get(i as usize) {
                    mapped
                } else {
                    i
                }
            } else if !source(c) {
                i
            } else if i == env {
                0
            } else if i < env {
                i + 1
            } else {
                i
            };
            local(&mut out, op, index);
        } else {
            out.extend_from_slice(
                bytes
                    .get(start..end)
                    .expect("native callable reader range exceeds input"),
            );
        }
    }
    out
}
/// Provider-private unset sentinel: origin=-1 with no function entry. A user
/// assigned null/undefined is never confused with an uninitialized binding.
pub fn is_unset(b: &mut Vec<u8>, p: &Plan, slot: u32) {
    local(b, 0x20, slot);
    gc(b, 0x1a);
    typed(b, 0x14, p.carrier);
    b.extend([0x04, TYPE_I32]);
    local(b, 0x20, slot);
    internal(b, p.carrier);
    field(b, 2, p.carrier, 2);
    i32c(b, -1);
    b.push(0x46);
    b.push(0x05);
    i32c(b, 0);
    b.push(0x0b);
}
pub fn initialize_aliases(b: &mut Vec<u8>, chunks: &[Chunk], p: &Plan, temp: u32) {
    use vybe_runtime::chunk::ConstExpr;
    let Some(root) = chunks.first() else {
        return;
    };
    let flag = p.cache_base + chunks.len() as u32;
    b.push(0x23);
    write_leb128_u32(b, flag);
    b.push(0x45);
    b.extend([0x04, TYPE_VOID]);
    let strings = super::sections::collect_string_constants(chunks);
    let hosts = super::sections::collect_host_globals(chunks);

    let mut seen = std::collections::BTreeSet::new();
    for c in chunks {
        for init in &c.global_inits {
            if let ConstExpr::RefFunc(index) = &init.init {
                if !seen.insert(init.name.clone()) {
                    continue;
                }
                let original = root.globals.iter().position(|n| n == &init.name);
                let g = original.unwrap_or_else(|| {
                    super::sections::rt_globals().len()
                        + strings.len()
                        + hosts
                            .iter()
                            .position(|n| n == &init.name)
                            .expect("native RefFunc alias imported")
                });
                let g_operand = if original.is_none() {
                    (g as u32) | 0x8000_0000
                } else {
                    g as u32
                };
                assert!(
                    *index < chunks.len(),
                    "initializer function target in module"
                );
                if hosts.contains(&init.name) {
                    b.push(0x23);
                    write_leb128_u32(b, g_operand);
                    local(b, 0x21, temp);
                    is_unset(b, p, temp);
                    b.extend([0x04, TYPE_VOID]);
                }
                b.push(0x23);
                write_leb128_u32(b, p.cache_base + *index as u32);
                external(b);
                b.push(0x24);
                write_leb128_u32(b, g_operand);
                if hosts.contains(&init.name) {
                    b.push(0x0b);
                }
            }
        }
    }
    i32c(b, 1);
    b.push(0x24);
    write_leb128_u32(b, flag);
    b.push(0x0b);
}
