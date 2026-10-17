//! Shared literal-byte membership spans. Language adapters prove byte storage and literal sets.
use vybe_runtime::chunk::{LocalName, LocalOrigin};
use vybe_runtime::{Chunk, Op};
pub fn emit_literal_set(
    chunks: &mut Vec<Chunk>,
    current: usize,
    take: bool,
    bytes: &[u8],
    line: u32,
) {
    assert!(!bytes.is_empty() && bytes.iter().all(|b| *b != 0));
    let name = format!("<primitive-memory32-span-set-{}-{:02x?}>", take, bytes);
    let target = if let Some(i) = chunks.iter().position(|c| c.name == name) {
        i
    } else {
        let mut c = Chunk::new(&name);
        c.arity = 1;
        c.param_count = 1;
        c.local_count = 1;
        let b = c.alloc_scratch(4);
        let (index, byte, matched, vector) = (b, b + 1, b + 2, b + 3);
        c.local_names.push(LocalName::new(
            vector,
            "<native-wasm-scratch:v128>",
            LocalOrigin::Compiler,
        ));
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::LOCAL_SET, index, 0);
        let (repeat, _) = c.emit_loop_s(0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, index, 0);
        c.emit_op(Op::I32_ADD, 0);
        // Membership is assigned only on the scalar edge path, so its slot
        // can retain the address through the vector guard and load.
        c.emit_op_u16(Op::LOCAL_TEE, matched, 0);
        c.emit_op(Op::I32_LOAD8_U, 0);
        c.emit_op_u16(Op::LOCAL_SET, byte, 0);
        // The byte load proves this page exists. A complete vector window also
        // checks that byte, so scalar membership is needed only at page edges.
        {
            c.emit_op_u16(Op::LOCAL_GET, matched, 0);
            c.emit_i32_const(65535, 0);
            c.emit_op(Op::I32_AND, 0);
            c.emit_i32_const(65520, 0);
            c.emit_op(Op::I32_LE_U, 0);
            c.emit_if(0);
            c.emit_op_u16(Op::LOCAL_GET, matched, 0);
            c.emit_op(Op::V128_LOAD, 0);
            c.emit_leb_u32(0x80, 0);
            c.emit_leb_u32(0, 0);
            c.emit_op_u16(Op::LOCAL_SET, vector, 0);
            if bytes.len() == 8 {
                let mut low = [0u8; 16];
                let mut high = [0u8; 16];
                for (i, value) in bytes.iter().enumerate() {
                    low[(value & 15) as usize] |= 1 << i;
                    high[(value >> 4) as usize] |= 1 << i;
                }
                c.emit_op(Op::V128_CONST, 0);
                for value in low {
                    c.emit(value, 0);
                }
                c.emit_op_u16(Op::LOCAL_GET, vector, 0);
                c.emit_op(Op::V128_CONST, 0);
                for _ in 0..16 {
                    c.emit(15, 0);
                }
                c.emit_op(Op::V128_AND, 0);
                c.emit_op(Op::I8X16_SWIZZLE, 0);
                c.emit_op(Op::V128_CONST, 0);
                for value in high {
                    c.emit(value, 0);
                }
                c.emit_op_u16(Op::LOCAL_GET, vector, 0);
                c.emit_i32_const(4, 0);
                c.emit_op(Op::I8X16_SHR_U, 0);
                c.emit_op(Op::I8X16_SWIZZLE, 0);
                c.emit_op(Op::V128_AND, 0);
                c.emit_op(Op::V128_CONST, 0);
                for _ in 0..16 {
                    c.emit(0, 0);
                }
                // Equality with zero is the nonmember mask. Taking members
                // uses it directly; skipping members needs its complement.
                c.emit_op(if take { Op::I8X16_EQ } else { Op::I8X16_NE }, 0);
                c.emit_op(Op::I8X16_BITMASK, 0);
            } else {
                for (i, value) in bytes.iter().enumerate() {
                    c.emit_op_u16(Op::LOCAL_GET, vector, 0);
                    c.emit_op(Op::V128_CONST, 0);
                    for _ in 0..16 {
                        c.emit(*value, 0);
                    }
                    c.emit_op(Op::I8X16_EQ, 0);
                    if i > 0 {
                        c.emit_op(Op::V128_OR, 0);
                    }
                }
                c.emit_op(Op::I8X16_BITMASK, 0);
            }
            if take && bytes.len() != 8 {
                c.emit_i32_const(65535, 0);
                c.emit_op(Op::I32_XOR, 0);
            }
            if !take {
                c.emit_op_u16(Op::LOCAL_GET, vector, 0);
                c.emit_op(Op::V128_CONST, 0);
                for _ in 0..16 {
                    c.emit(0, 0);
                }
                c.emit_op(Op::I8X16_EQ, 0);
                c.emit_op(Op::I8X16_BITMASK, 0);
                c.emit_op(Op::I32_OR, 0);
            }
            c.emit_op_u16(Op::LOCAL_TEE, byte, 0);
            c.emit_if(0);
            c.emit_op_u16(Op::LOCAL_GET, index, 0);
            c.emit_op_u16(Op::LOCAL_GET, byte, 0);
            c.emit_op(Op::I32_CTZ, 0);
            c.emit_op(Op::I32_ADD, 0);
            c.emit_op(Op::F64_CONVERT_I32_U, 0);
            c.emit_op(Op::RETURN, 0);
            c.emit_end(0);
            c.emit_op_u16(Op::LOCAL_GET, index, 0);
            c.emit_i32_const(16, 0);
            c.emit_op(Op::I32_ADD, 0);
            c.emit_op_u16(Op::LOCAL_SET, index, 0);
            c.emit_br(1, 0);
            c.emit_end(0);
        }
        c.emit_op_u16(Op::LOCAL_GET, byte, 0);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_if(0);
        c.emit_op_u16(Op::LOCAL_GET, index, 0);
        c.emit_op(Op::F64_CONVERT_I32_U, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
        let mut membership_branches = Vec::new();
        for value in bytes {
            c.emit_op_u16(Op::LOCAL_GET, byte, 0);
            c.emit_i32_const(*value as i32, 0);
            c.emit_op(Op::I32_EQ, 0);
            membership_branches.push(c.emit_if_value(0));
            c.emit_i32_const(1, 0);
            c.emit_else(0);
        }
        c.emit_i32_const(0, 0);
        for branch in membership_branches.into_iter().rev() {
            c.emit_end(0);
            c.patch_block(branch);
        }
        crate::primitives::instructions::core_wasm::local_tee(&mut c, 0, matched);
        if take {
            c.emit_op(Op::I32_EQZ, 0);
        }
        c.emit_if(0);
        c.emit_op_u16(Op::LOCAL_GET, index, 0);
        c.emit_op(Op::F64_CONVERT_I32_U, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
        c.emit_op_u16(Op::LOCAL_GET, index, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_op_u16(Op::LOCAL_SET, index, 0);
        c.emit_br(0, 0);
        c.emit_end(0);
        c.patch_loop(repeat);
        c.emit_f64_const(0.0, 0);
        c.emit_op(Op::RETURN, 0);

        c.name = format!("{name}-body").into();
        c.memory_min_pages = chunks[0].memory_min_pages.clone();
        c.memory_is_64 = chunks[0].memory_is_64.clone();
        let body_target = chunks.len();
        let helper = c.clone();
        chunks.push(c);
        let mut caller = Chunk::new("literal_span");
        caller.arity = 1;
        caller.param_count = 1;
        caller.local_count = 1;
        let b = caller.alloc_scratch(2);
        let (byte, matched) = (b, b + 1);
        for offset in 0..2 {
            caller.emit_op_u16(Op::LOCAL_GET, 0, 0);
            if offset != 0 {
                caller.emit_i32_const(offset, 0);
                caller.emit_op(Op::I32_ADD, 0);
            }
            caller.emit_op(Op::I32_LOAD8_U, 0);
            crate::primitives::instructions::core_wasm::local_tee(&mut caller, 0, byte);
            caller.emit_op(Op::I32_EQZ, 0);
            caller.emit_if(0);
            caller.emit_f64_const(offset as f64, 0);
            caller.emit_op(Op::RETURN, 0);
            caller.emit_end(0);
            let mut branches = Vec::new();
            for value in bytes {
                caller.emit_op_u16(Op::LOCAL_GET, byte, 0);
                caller.emit_i32_const(*value as i32, 0);
                caller.emit_op(Op::I32_EQ, 0);
                branches.push(caller.emit_if_value(0));
                caller.emit_i32_const(1, 0);
                caller.emit_else(0);
            }
            caller.emit_i32_const(0, 0);
            for branch in branches.into_iter().rev() {
                caller.emit_end(0);
                caller.patch_block(branch);
            }
            crate::primitives::instructions::core_wasm::local_tee(&mut caller, 0, matched);
            if take {
                caller.emit_op(Op::I32_EQZ, 0);
            }
            caller.emit_if(0);
            caller.emit_f64_const(offset as f64, 0);
            caller.emit_op(Op::RETURN, 0);
            caller.emit_end(0);
        }
        caller.emit_op_u16(Op::LOCAL_GET, 0, 0);
        caller.emit_i32_const(2, 0);
        caller.emit_op(Op::I32_ADD, 0);
        caller
            .emit_direct_defined_call(body_target as u32, &helper, false, 0)
            .unwrap();
        caller.emit_f64_const(2.0, 0);
        caller.emit_op(Op::F64_ADD, 0);
        caller.emit_op(Op::RETURN, 0);

        caller.name = name.into();
        caller.memory_min_pages = chunks[0].memory_min_pages.clone();
        caller.memory_is_64 = chunks[0].memory_is_64.clone();
        let target = chunks.len();
        chunks.push(caller);
        target
    };
    if super::functions::try_emit_direct_helper_call(chunks, current, target, 1, line) {
        return;
    }
    let helper = chunks[target].clone();
    chunks[current]
        .emit_direct_defined_call(target as u32, &helper, false, line)
        .expect("compiler-owned literal-set span");
}
