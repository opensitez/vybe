use vybe_compiler::primitives::{self, packing};
use vybe_runtime::{Chunk, Op, Value, VM};

fn program(width: usize, endian: packing::Endian) -> Vec<Chunk> {
    let mut root = Chunk::new("packing-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut packer = Chunk::new("packer");
    packer.arity = 1;
    packer.param_count = 1;
    packer.local_count = 1;
    let mut chunks = vec![root, packer];
    if width == 2 {
        packing::emit_pack_u16_from_f64_slot(&mut chunks, 1, 0, endian, 0);
    } else {
        packing::emit_pack_u32_from_f64_slot(&mut chunks, 1, 0, endian, 0);
    }
    assert_eq!(chunks[1].local_count, 2);
    chunks[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
    primitives::globals::declare_free_globals(&mut chunks);
    primitives::globals::normalize_global_table(&mut chunks);
    chunks
}

#[test]
fn packing_preserves_endian_bytes_high_bits_and_numeric_traps() {
    for width in [2, 4] {
        for little in [false, true] {
            let endian = if little { packing::Endian::Little } else { packing::Endian::Big };
            let chunks = program(width, endian);
            let native = vybe_platform_wasm::write_wasm(&chunks);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&native).unwrap();
            let mut vm = VM::new();
            primitives::platforms::register_platforms_all(&mut vm);
            let function = vm.run(chunks).unwrap();
            let mut seed = 0x12345678u32;
            for n in (0..65536u32).chain((0..4096).map(|_| {
                seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
                seed
            })) {
                let result = vm.invoke(&function, &[Value::F64(n as f64 + 0.75)]).unwrap();
                let expected: Vec<u16> = (0..width).map(|i| {
                    let byte = if little { i } else { width - i - 1 };
                    ((n >> (8 * byte)) & 255) as u16
                }).collect();
                assert_eq!(result.as_str().encode_utf16().collect::<Vec<_>>(), expected);
                assert!(vm.last_exception.is_none());
            }
            for input in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY, -1.0, 4294967296.0] {
                let mut vm = VM::new();
                primitives::platforms::register_platforms_all(&mut vm);
                let function = vm.run(program(width, endian)).unwrap();
                assert!(vm.invoke(&function, &[Value::F64(input)]).is_err()
                    || vm.last_exception.is_some());
            }
        }
    }
}

#[test]
fn unsigned_packing_keeps_one_scratch_at_the_local_index_boundary() {
    for width in [2, 4] {
        let mut chunk = Chunk::new("packing-local-boundary");
        chunk.local_count = u16::MAX - 1;
        let mut chunks = vec![chunk];
        if width == 2 {
            packing::emit_pack_u16_from_f64_slot(&mut chunks, 0, 0, packing::Endian::Little, 0);
        } else {
            packing::emit_pack_u32_from_f64_slot(&mut chunks, 0, 0, packing::Endian::Little, 0);
        }
        assert_eq!(chunks[0].local_count, u16::MAX);
    }
}
