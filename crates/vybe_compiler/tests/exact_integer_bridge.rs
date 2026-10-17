use std::sync::Arc;
use vybe_compiler::primitives::{exact_integer, functions, platforms};
use vybe_runtime::{Chunk, Op, VM, Value, bigint::BigIntVal};

const WORDS: &[u64] = &[
    0,
    1,
    u32::MAX as u64,
    1u64 << 32,
    (1u64 << 53) - 1,
    1u64 << 53,
    (1u64 << 53) + 1,
    (-9_007_199_254_740_993i64) as u64,
    0x0123_4567_89ab_cdef,
    0xfedc_ba98_7654_3210,
    i64::MAX as u64,
    i64::MIN as u64,
    u64::MAX,
    u64::MAX - 1,
];

fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
            let mut reader = body.get_operators_reader().unwrap();
            while !reader.eof() {
                assert!(
                    !matches!(
                        reader.read().unwrap(),
                        wasmparser::Operator::F64ConvertI64S | wasmparser::Operator::F64ConvertI64U
                    ),
                    "whole-word floating conversion in exact bridge"
                );
            }
        }
    }
}

fn runtime() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}

fn assert_bigint(value: Value, expected: &str) {
    let Value::BigInt(value) = value else {
        panic!("expected exact BigInt carrier, got {value:?}")
    };
    assert_eq!(value.to_string(), expected);
}

fn push_word(chunks: &mut [Chunk], word: u64) {
    chunks[0].emit_i64_const(word as i64, 0);
    exact_integer::emit_box_i64(chunks, 0, false, 0);
}

#[test]
fn exact_binary_operations_use_native_word_semantics() {
    use vybe_ast::IntOp;
    for (operation, left, right, expected) in [
        (IntOp::Add, u64::MAX, 2, 1),
        (IntOp::Sub, 0, 1, u64::MAX),
        (IntOp::Mul, (1u64 << 63) + 1, 3, (1u64 << 63) + 3),
        (IntOp::DivU, u64::MAX, 3, u64::MAX / 3),
        (IntOp::DivS, (-7i64) as u64, 3, (-2i64) as u64),
        (IntOp::RemS, (-7i64) as u64, 3, u64::MAX),
        (IntOp::RemU, u64::MAX, 7, u64::MAX % 7),
        (IntOp::Shl, 1, 65, 2),
        (IntOp::ShrU, u64::MAX, 63, 1),
        (IntOp::ShrS, (1u64 << 63), 63, u64::MAX),
        (
            IntOp::And,
            u64::MAX,
            0x0123_4567_89ab_cdef,
            0x0123_4567_89ab_cdef,
        ),
        (IntOp::Or, 1u64 << 63, 1, (1u64 << 63) + 1),
        (IntOp::Xor, u64::MAX, 1, u64::MAX - 1),
    ] {
        let mut chunks = vec![Chunk::new("exact-binary-word")];
        push_word(&mut chunks, left);
        push_word(&mut chunks, right);
        exact_integer::emit_integer_binary(&mut chunks, 0, operation, false, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        validate(&chunks);
        assert_bigint(runtime().run(chunks).unwrap(), &expected.to_string());
    }
}

#[test]
fn exact_comparisons_return_native_predicates_not_dynamic_equality() {
    use vybe_ast::IntOp;
    for (operation, expected) in [
        (IntOp::Eq, 0),
        (IntOp::Ne, 1),
        (IntOp::LtS, 1),
        (IntOp::LtU, 0),
        (IntOp::LeS, 1),
        (IntOp::LeU, 0),
        (IntOp::GtS, 0),
        (IntOp::GtU, 1),
        (IntOp::GeS, 0),
        (IntOp::GeU, 1),
    ] {
        let mut chunks = vec![Chunk::new("exact-word-comparison")];
        push_word(&mut chunks, u64::MAX);
        push_word(&mut chunks, 1);
        exact_integer::emit_integer_binary(&mut chunks, 0, operation, true, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        validate(&chunks);
        assert!(!vybe_runtime::debug::disassemble(&chunks[0]).contains("dyneq"));
        assert_eq!(runtime().run(chunks).unwrap(), Value::I32(expected));
    }
}

#[test]
fn nested_exact_operations_preserve_outer_values() {
    let mut chunks = vec![Chunk::new("nested-exact-word")];
    push_word(&mut chunks, (1u64 << 53) + 1);
    push_word(&mut chunks, 1);
    push_word(&mut chunks, 2);
    exact_integer::emit_integer_binary(&mut chunks, 0, vybe_ast::IntOp::Add, false, 0);
    exact_integer::emit_integer_binary(&mut chunks, 0, vybe_ast::IntOp::Mul, false, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    validate(&chunks);
    assert_bigint(
        runtime().run(chunks).unwrap(),
        &(((1u64 << 53) + 1) * 3).to_string(),
    );
}

#[test]
fn normalization_preserves_bigints_and_truncates_only_numeric_inputs() {
    for signed in [false, true] {
        let mut root = Chunk::new("exact-normalize-owner");
        functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut body = Chunk::new("exact-normalize-input");
        body.arity = 1;
        body.param_count = 1;
        body.local_count = 1;
        body.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let mut chunks = vec![root, body];
        exact_integer::emit_normalize_value(&mut chunks, 1, signed, 0);
        chunks[1].emit_op(Op::RETURN, 0);
        validate(&chunks);
        let mut vm = runtime();
        let callable = vm.run(chunks).unwrap();
        for (input, signed_expected, unsigned_expected) in [
            (Value::F64(3.75), "3", "3"),
            (Value::F64(-3.75), "-3", "18446744073709551613"),
            (Value::I32(-1), "-1", "18446744073709551615"),
            (
                Value::BigInt(Arc::new(BigIntVal::parse("9007199254740993").unwrap())),
                "9007199254740993",
                "9007199254740993",
            ),
            (
                Value::BigInt(Arc::new(BigIntVal::parse("18446744073709551615").unwrap())),
                "-1",
                "18446744073709551615",
            ),
        ] {
            assert_bigint(
                vm.invoke(&callable, &[input]).unwrap(),
                if signed {
                    signed_expected
                } else {
                    unsigned_expected
                },
            );
        }
    }
}

#[test]
fn native_words_box_with_exact_signed_and_unsigned_readings() {
    for &word in WORDS {
        for signed in [false, true] {
            let mut chunks = vec![Chunk::new("exact-word-box")];
            chunks[0].emit_i64_const(word as i64, 0);
            exact_integer::emit_box_i64(&mut chunks, 0, signed, 0);
            chunks[0].emit_op(Op::RETURN, 0);
            validate(&chunks);
            let expected = if signed {
                (word as i64).to_string()
            } else {
                word.to_string()
            };
            assert_bigint(runtime().run(chunks).unwrap(), &expected);
        }
    }
}

#[test]
fn value_memory_word_bridges_preserve_declared_transport_and_native_bits() {
    use vybe_ast::WideIntegerTransport;
    use vybe_compiler::primitives::pointers;
    for transport in [WideIntegerTransport::Number, WideIntegerTransport::Exact] {
        for signed in [false, true] {
            let words: Vec<u64> = if transport == WideIntegerTransport::Exact {
                WORDS.to_vec()
            } else if signed {
                vec![0, 5, (-3i64) as u64, i64::MIN as u64]
            } else {
                vec![0, 5, (-3i64) as u64, 1u64 << 63]
            };
            for word in words {
                let mut root = Chunk::new("value-word-owner");
                root.memory_min_pages = vec![1];
                functions::emit_ref_func(&mut root, 1, 0, 0);
                root.emit_op(Op::RETURN, 0);
                let mut leaf = Chunk::new("value-word-memory");
                leaf.arity = 1;
                leaf.param_count = 1;
                leaf.local_count = 1;
                leaf.emit_i32_const(7, 0);
                leaf.emit_op_u16(Op::LOCAL_GET, 0, 0);
                let mut chunks = vec![root, leaf];
                pointers::emit_linear_i64_value_store(&mut chunks, 1, signed, transport, 0);
                chunks[1].emit_op(Op::DROP, 0);
                chunks[1].emit_i32_const(7, 0);
                pointers::emit_linear_i64_value_load(&mut chunks, 1, signed, transport, 0);
                chunks[1].emit_op(Op::RETURN, 0);
                let bytes = vybe_platform_wasm::write_wasm_with_options(
                    &chunks,
                    vybe_platform_wasm::WasmWriteOptions {
                        include_vybe_metadata: false,
                        include_name_section: true,
                    },
                );
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&bytes)
                    .unwrap();
                let mut vm = runtime();
                let callable = vm.run(chunks).unwrap();
                let input = if transport == WideIntegerTransport::Exact {
                    Value::BigInt(Arc::new(BigIntVal::parse(&word.to_string()).unwrap()))
                } else if signed || word == (-3i64) as u64 {
                    Value::F64((word as i64) as f64)
                } else {
                    Value::F64(word as f64)
                };
                let result = vm.invoke(&callable, &[input]).unwrap();
                if transport == WideIntegerTransport::Exact {
                    assert_bigint(
                        result,
                        &if signed {
                            (word as i64).to_string()
                        } else {
                            word.to_string()
                        },
                    );
                } else {
                    assert_eq!(
                        result.as_f64(),
                        if signed {
                            (word as i64) as f64
                        } else {
                            word as f64
                        }
                    );
                }
                vm.memory
                    .with_buffer(|memory| assert_eq!(&memory[7..15], &word.to_le_bytes()));
            }
        }
    }
}

#[test]
fn box_unbox_roundtrips_keep_every_byte_in_unaligned_memory() {
    for &word in WORDS {
        for signed in [false, true] {
            let mut chunks = vec![Chunk::new("exact-word-roundtrip")];
            chunks[0].memory_min_pages = vec![1];
            chunks[0].emit_i32_const(7, 0);
            chunks[0].emit_i64_const(word as i64, 0);
            exact_integer::emit_box_i64(&mut chunks, 0, signed, 0);
            exact_integer::emit_unbox_i64(&mut chunks, 0, 0);
            chunks[0].emit_op(Op::I64_STORE, 0);
            chunks[0].emit_i32_const(7, 0);
            chunks[0].emit_op(Op::I64_LOAD, 0);
            exact_integer::emit_box_i64(&mut chunks, 0, signed, 0);
            chunks[0].emit_op(Op::RETURN, 0);
            validate(&chunks);
            let mut vm = runtime();
            let expected = if signed {
                (word as i64).to_string()
            } else {
                word.to_string()
            };
            assert_bigint(vm.run(chunks).unwrap(), &expected);
            vm.memory
                .with_buffer(|memory| assert_eq!(&memory[7..15], &word.to_le_bytes()));
        }
    }
}

#[test]
fn independent_bigint_inputs_unbox_modulo_64_without_whole_word_numbers() {
    let mut root = Chunk::new("exact-unbox-owner");
    root.memory_min_pages = vec![1];
    functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut body = Chunk::new("exact-unbox-input");
    body.arity = 1;
    body.param_count = 1;
    body.local_count = 1;
    body.emit_i32_const(9, 0);
    body.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut chunks = vec![root, body];
    exact_integer::emit_unbox_i64(&mut chunks, 1, 0);
    chunks[1].emit_op(Op::I64_STORE, 0);
    chunks[1].emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunks[1].emit_op(Op::RETURN, 0);
    validate(&chunks);
    let mut vm = runtime();
    let callable = vm.run(chunks).unwrap();
    for (decimal, word) in [
        ("9007199254740993", (1u64 << 53) + 1),
        ("9223372036854775807", i64::MAX as u64),
        ("-9223372036854775808", i64::MIN as u64),
        ("18446744073709551615", u64::MAX),
        ("18446744073709551617", 1),
        ("-18446744073709551617", u64::MAX),
    ] {
        let value = Value::BigInt(Arc::new(BigIntVal::parse(decimal).unwrap()));
        assert_bigint(vm.invoke(&callable, &[value]).unwrap(), decimal);
        vm.memory
            .with_buffer(|memory| assert_eq!(&memory[9..17], &word.to_le_bytes()));
    }
}

#[test]
fn bridge_preserves_live_raw_operands_singleton_and_reclaimed_value_locals() {
    let left = 0xfedc_ba98_7654_3210u64;
    let right = 0x0123_4567_89ab_cdefu64;
    let sentinel = 0xa5a5_5a5a_1234_5678u64;
    let mut root = Chunk::new("exact-word-scratch-isolation");
    root.memory_min_pages = vec![1];
    let value = root.alloc_scratch(4);
    root.emit_i32_const(73, 0);
    root.emit_op_u16(Op::LOCAL_SET, value, 0);
    let singleton = root.i64_scratch();
    root.emit_i64_const(sentinel as i64, 0);
    root.emit_op_u16(Op::LOCAL_SET, singleton, 0);
    let prior_peak = root.scratch_high_water;
    root.local_count = 0;
    root.emit_i32_const(7, 0);
    root.emit_i64_const(left as i64, 0);
    root.emit_i64_const(right as i64, 0);
    let mut chunks = vec![root];
    exact_integer::emit_box_i64(&mut chunks, 0, false, 0);
    exact_integer::emit_unbox_i64(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::I64_SUB, 0);
    chunks[0].emit_op(Op::I64_STORE, 0);
    chunks[0].emit_i32_const(23, 0);
    chunks[0].emit_op_u16(Op::LOCAL_GET, singleton, 0);
    chunks[0].emit_op(Op::I64_STORE, 0);
    chunks[0].emit_op_u16(Op::LOCAL_GET, value, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    assert_eq!(chunks[0].i64_scratch_slot, Some(singleton));
    assert!(chunks[0].i64_locals.iter().any(|&slot| slot >= prior_peak));
    assert!(!chunks[0].i64_locals.contains(&value));
    validate(&chunks);
    let mut vm = runtime();
    assert_eq!(vm.run(chunks).unwrap(), Value::I32(73));
    vm.memory.with_buffer(|memory| {
        assert_eq!(&memory[7..15], &left.wrapping_sub(right).to_le_bytes());
        assert_eq!(&memory[23..31], &sentinel.to_le_bytes());
    });
}

#[test]
fn two_boxed_operands_unbox_without_clobbering_the_saved_right_word() {
    let left = 0xfedc_ba98_7654_3210u64;
    let right = 0x0123_4567_89ab_cdefu64;
    let mut root = Chunk::new("exact-word-nested-operands");
    root.memory_min_pages = vec![1];
    root.emit_i32_const(7, 0);
    root.emit_i64_const(left as i64, 0);
    let mut chunks = vec![root];
    exact_integer::emit_box_i64(&mut chunks, 0, true, 0);
    chunks[0].emit_i64_const(right as i64, 0);
    exact_integer::emit_box_i64(&mut chunks, 0, false, 0);
    exact_integer::emit_unbox_i64(&mut chunks, 0, 0);
    let saved_right = chunks[0].alloc_scratch(1);
    chunks[0].i64_locals.insert(saved_right);
    chunks[0].emit_op_u16(Op::LOCAL_SET, saved_right, 0);
    exact_integer::emit_unbox_i64(&mut chunks, 0, 0);
    chunks[0].emit_op_u16(Op::LOCAL_GET, saved_right, 0);
    chunks[0].emit_op(Op::I64_SUB, 0);
    chunks[0].emit_op(Op::I64_STORE, 0);
    chunks[0].emit_i32_const(7, 0);
    chunks[0].emit_op(Op::I64_LOAD, 0);
    exact_integer::emit_box_i64(&mut chunks, 0, false, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    validate(&chunks);
    let mut vm = runtime();
    let expected = left.wrapping_sub(right);
    assert_bigint(vm.run(chunks).unwrap(), &expected.to_string());
    vm.memory
        .with_buffer(|memory| assert_eq!(&memory[7..15], &expected.to_le_bytes()));
}

#[test]
fn exact_boxed_values_join_both_backing_branches() {
    let words = [0x0123_4567_89ab_cdefu64, u64::MAX];
    for choose_first in [false, true] {
        let mut chunks = vec![Chunk::new("exact-word-value-join")];
        chunks[0].emit_i32_const(i32::from(choose_first), 0);
        chunks[0].emit_if_value(0);
        chunks[0].emit_i64_const(words[0] as i64, 0);
        exact_integer::emit_box_i64(&mut chunks, 0, false, 0);
        chunks[0].emit_else(0);
        chunks[0].emit_i64_const(words[1] as i64, 0);
        exact_integer::emit_box_i64(&mut chunks, 0, false, 0);
        chunks[0].emit_end(0);
        chunks[0].emit_op(Op::RETURN, 0);
        validate(&chunks);
        assert_bigint(
            runtime().run(chunks).unwrap(),
            &words[usize::from(!choose_first)].to_string(),
        );
    }
}
