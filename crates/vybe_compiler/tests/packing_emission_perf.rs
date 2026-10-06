//! Packing conversion reuse, endian byte layout, and binary validation.
use vybe_compiler::primitives::{bundle, globals, packing, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};
use vybex as _;
#[path = "fixtures/repeated_packing.rs"]
mod baseline;
#[path = "fixtures/repeated_php_packing.rs"]
mod php_baseline;
fn vm(input: f64) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", Value::F64(input));
    vm
}
fn program(old: bool, width: usize, little: bool, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<packing>")];
    chunks[0].local_count = 1;
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "input", 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
        if old {
            let endian = if little {
                baseline::Endian::Little
            } else {
                baseline::Endian::Big
            };
            match width {
                1 => baseline::emit_pack_byte_from_f64_slot(&mut chunks, 0, 0, 0),
                2 => baseline::emit_pack_u16_from_f64_slot(&mut chunks, 0, 0, endian, 0),
                4 => baseline::emit_pack_u32_from_f64_slot(&mut chunks, 0, 0, endian, 0),
                _ => unreachable!(),
            }
        } else {
            let endian = if little {
                packing::Endian::Little
            } else {
                packing::Endian::Big
            };
            match width {
                1 => packing::emit_pack_byte_from_f64_slot(&mut chunks, 0, 0, 0),
                2 => packing::emit_pack_u16_from_f64_slot(&mut chunks, 0, 0, endian, 0),
                4 => packing::emit_pack_u32_from_f64_slot(&mut chunks, 0, 0, endian, 0),
                _ => unreachable!(),
            }
        }
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
#[test]
fn packing_keeps_endian_layout_and_high_unsigned_bits() {
    for width in [2, 4] {
        for little in [false, true] {
            let old = program(true, width, little, 1);
            let new = program(false, width, little, 1);
            validate(&old);
            validate(&new);
            for value in [
                0.0,
                1.0,
                255.0,
                256.0,
                4660.0,
                65535.0,
                65536.0,
                2147483647.0,
                2147483648.0,
                4294967295.0,
                1.75,
            ] {
                let bits = value as u32;
                let bytes = if little {
                    bits.to_le_bytes()[..width].to_vec()
                } else {
                    bits.to_be_bytes()[4 - width..].to_vec()
                };
                let expected: String = bytes.into_iter().map(char::from).collect();
                assert_eq!(vm(value).run(old.clone()).unwrap().as_str(), expected);
                assert_eq!(vm(value).run(new.clone()).unwrap().as_str(), expected);
            }
        }
    }
}
#[test]
fn single_byte_signed_adjustment_validates_and_keeps_existing_layout() {
    let old = program(true, 1, true, 1);
    let new = program(false, 1, true, 1);
    validate(&new);
    for value in [-256.0, -255.0, -128.0, -1.0, 0.0, 255.0, 256.0, 511.0] {
        assert_eq!(
            vm(value).run(new.clone()).unwrap().as_str(),
            vm(value).run(old.clone()).unwrap().as_str()
        );
    }
}
#[test]
fn packing_emits_one_conversion_and_uses_one_private_slot() {
    for width in [2, 4] {
        for little in [false, true] {
            let old = program(true, width, little, 20);
            let new = program(false, width, little, 20);
            let disassembly = vybe_runtime::debug::disassemble(&new[0]);
            assert_eq!(
                disassembly.matches("i32.trunc_f64_u").count(),
                20,
                "{disassembly}"
            );
            assert_eq!(new[0].local_count, 21);
            validate(&new);
            let before: usize = old.iter().map(|c| c.code.len()).sum();
            let after: usize = new.iter().map(|c| c.code.len()).sum();
            println!(
                "packing width={width} little={little}: bytecode={before}/{after} locals={}/{} wasm={}/{}",
                old[0].local_count,
                new[0].local_count,
                vybe_platform_wasm::write_wasm(&old).len(),
                vybe_platform_wasm::write_wasm(&new).len()
            );
            assert!(after < before);
            assert_eq!(
                vm(305419896.0).run(new).unwrap().as_str(),
                vm(305419896.0).run(old).unwrap().as_str()
            );
        }
    }
}
#[test]
fn php_packing_frontend_preserves_supported_unsigned_layout() {
    vybe_language_php::register();
    for (body, expected) in [
        ("return bin2hex(pack('C', 255));", "ff"),
        ("return bin2hex(pack('C', -1.5));", "ff"),
        ("return bin2hex(pack('C', -257));", "ff"),
        ("return bin2hex(pack('n', 4660));", "1234"),
        ("return bin2hex(pack('v', 4660));", "3412"),
        ("return bin2hex(pack('N', 305419896));", "12345678"),
        ("return bin2hex(pack('V', 305419896));", "78563412"),
    ] {
        let module = vybe_language_php::parse(&format!("<?php {body}")).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        validate(&chunks);
        assert_eq!(vm(0.0).run(chunks).unwrap().as_str(), expected, "{body}");
    }
}
#[test]
fn php_integer_packing_wraps_before_checked_numeric_emission() {
    vybe_language_php::register();
    for (format, input, expected) in [
        ("n", "-1", "ffff"),
        ("v", "-1.5", "ffff"),
        ("N", "-257", "fffffeff"),
        ("V", "-257", "fffeffff"),
        ("N", "4294967297", "00000001"),
        ("N", "9223372036854775807", "ffffffff"),
        ("V", "-9223372036854775807", "01000000"),
        ("C", "9223372036854775807", "ff"),
        ("N", "'12junk'", "0000000c"),
        ("n", "'-1'", "ffff"),
        ("V", "true", "01000000"),
        ("N", "null", "00000000"),
        ("N", "INF", "00000000"),
        ("N", "NAN", "00000000"),
        ("CCC", "-1, -257, 4294967297", "ffff01"),
        ("C2", "-1.5, -0.5", "ff00"),
    ] {
        let source = format!(
            "<?php $format = '{format}'; $value = {input}; return bin2hex(pack($format, $value));"
        );
        // Multi-value formats supply ordinary separate arguments.
        let source = if input.contains(',') {
            format!("<?php $format = '{format}'; return bin2hex(pack($format, {input}));")
        } else {
            source
        };
        let module = vybe_language_php::parse(&source).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        validate(&chunks);
        assert_eq!(vm(0.0).run(chunks).unwrap().as_str(), expected, "{source}");
    }
}

#[test]
#[ignore = "manual interleaved debug packing benchmark"]
fn packing_execution_benchmark() {
    for width in [2, 4] {
        let old = program(true, width, true, 1024);
        let new = program(false, width, true, 1024);
        let mut times = [Vec::new(), Vec::new()];
        for repeat in 0..5 {
            for which in if repeat % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut vm = vm(305419896.0);
                let chunks = if which == 0 { old.clone() } else { new.clone() };
                let start = std::time::Instant::now();
                let result = vm.run(chunks).unwrap();
                times[which].push(start.elapsed().as_micros());
                assert_eq!(result.as_str().len(), if width == 2 { 2 } else { 4 });
            }
        }
        for times in &mut times {
            times.sort();
        }
        println!(
            "packing width={width}: repeated={}us cached={}us speedup={:.2}x",
            times[0][2],
            times[1][2],
            times[0][2] as f64 / times[1][2] as f64
        );
    }
}

#[test]
fn byte_packing_truncates_before_wrapping_finite_values() {
    let chunks = program(false, 1, true, 1);
    validate(&chunks);
    for (value, byte) in [
        (-257.0, 255u8),
        (-1.5, 255),
        (-0.5, 0),
        (1.5, 1),
        (4294967297.0, 1),
    ] {
        assert_eq!(
            vm(value).run(chunks.clone()).unwrap().as_str(),
            char::from(byte).to_string()
        );
    }
}
#[test]
#[ignore = "manual export for independent Wasm execution and validation"]
fn export_packing_parameter_fixtures() {
    let dir = std::env::temp_dir().join("vybe-packing-wasm");
    std::fs::create_dir_all(&dir).unwrap();
    for width in [1, 2, 4] {
        for endian in [packing::Endian::Little, packing::Endian::Big] {
            let mut chunks = vec![Chunk::new("<packing-parameter>")];
            chunks[0].arity = 1;
            chunks[0].local_count = 1;
            match width {
                1 => packing::emit_pack_byte_from_f64_slot(&mut chunks, 0, 0, 0),
                2 => packing::emit_pack_u16_from_f64_slot(&mut chunks, 0, 0, endian, 0),
                4 => packing::emit_pack_u32_from_f64_slot(&mut chunks, 0, 0, endian, 0),
                _ => unreachable!(),
            }
            chunks[0].emit_op(Op::RETURN, 0);
            validate(&chunks);
            std::fs::write(
                dir.join(format!("width-{width}-{endian:?}.wasm")),
                vybe_platform_wasm::write_wasm(&chunks),
            )
            .unwrap();
        }
    }
}

fn php_pack_program(old: bool, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<php-pack>")];
    for call in 0..calls {
        chunks[0].emit_string_const("N", 0);
        chunks[0].emit_f64_const(305419896.0, 0);
        if old {
            php_baseline::emit_php_pack(&mut chunks, 0, 2, 0);
        } else {
            vybe_language_php::emitter::packing_adapter::emit_php_pack(&mut chunks, 0, 2, 0);
        }
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
#[test]
fn php_pack_predicates_avoid_redundant_truthiness_dispatch() {
    let old = php_pack_program(true, 20);
    let new = php_pack_program(false, 20);
    validate(&old);
    validate(&new);
    let before: usize = old.iter().map(|c| c.code.len()).sum();
    let after: usize = new.iter().map(|c| c.code.len()).sum();
    println!(
        "PHP pack 20 calls: bytecode={before}/{after} wasm={}/{}",
        vybe_platform_wasm::write_wasm(&old).len(),
        vybe_platform_wasm::write_wasm(&new).len()
    );
    assert!(after < before);
    assert_eq!(
        vm(0.0).run(old).unwrap().as_str(),
        vm(0.0).run(new).unwrap().as_str()
    );
}
#[test]
#[ignore = "manual PHP packing interleaved diagnostic benchmark"]
fn php_pack_execution_benchmark() {
    let programs = [php_pack_program(true, 128), php_pack_program(false, 128)];
    let mut times = [Vec::new(), Vec::new()];
    for repeat in 0..5 {
        for which in if repeat % 2 == 0 { [0, 1] } else { [1, 0] } {
            let mut vm = vm(0.0);
            let chunks = programs[which].clone();
            let start = std::time::Instant::now();
            assert_eq!(vm.run(chunks).unwrap().as_str(), "\u{12}4Vx");
            times[which].push(start.elapsed().as_micros());
        }
    }
    times[0].sort();
    times[1].sort();
    println!(
        "PHP pack diagnostic: before={}us after={}us speedup={:.2}x",
        times[0][2],
        times[1][2],
        times[0][2] as f64 / times[1][2] as f64
    );
}
#[test]
#[ignore = "manual PHP pack exports for native Wasm execution"]
fn export_php_packing_parameter_fixtures() {
    let dir = std::env::temp_dir().join("vybe-packing-wasm");
    std::fs::create_dir_all(&dir).unwrap();
    for format in ["C", "n", "v", "N", "V"] {
        let mut chunks = vec![Chunk::new("<php-pack-parameter>")];
        chunks[0].arity = 1;
        chunks[0].local_count = 1;
        chunks[0].emit_string_const(format, 0);
        chunks[0].emit_op_u16(Op::LOCAL_GET, 0, 0);
        vybe_language_php::emitter::packing_adapter::emit_php_pack(&mut chunks, 0, 2, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        bundle::finalize_with_runtime_helpers(&mut chunks);
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        validate(&chunks);
        std::fs::write(
            dir.join(format!("php-format-{:02x}.wasm", format.as_bytes()[0])),
            vybe_platform_wasm::write_wasm(&chunks),
        )
        .unwrap();
    }
    for format in ["N", "V", "n", "v", "Cfirst/Csecond", "C*"] {
        let mut chunks = vec![Chunk::new("<php-unpack-parameter>")];
        chunks[0].arity = 1;
        chunks[0].local_count = 1;
        chunks[0].emit_string_const(format, 0);
        chunks[0].emit_op_u16(Op::LOCAL_GET, 0, 0);
        vybe_language_php::emitter::packing_adapter::emit_php_unpack(&mut chunks, 0, 2, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        bundle::finalize_with_runtime_helpers(&mut chunks);
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        validate(&chunks);
        let name = format
            .as_bytes()
            .iter()
            .map(|b| format!("{b:02x}"))
            .collect::<String>();
        std::fs::write(
            dir.join(format!("php-unpack-{name}.wasm")),
            vybe_platform_wasm::write_wasm(&chunks),
        )
        .unwrap();
    }
}

#[test]
fn php_format_guard_preserves_dynamic_fallback() {
    for format in [
        Value::Null,
        Value::Bool(false),
        Value::Bool(true),
        Value::F64(0.0),
        Value::F64(1.0),
        Value::String("N".into()),
        Value::String("n".into()),
        Value::String("V".into()),
        Value::String("v".into()),
    ] {
        let build = |old| {
            let mut chunks = vec![Chunk::new("<php-format-guard>")];
            vybe_compiler::primitives::instructions::core_wasm::emit_value(
                &mut chunks[0],
                0,
                &format,
            );
            chunks[0].emit_f64_const(305419896.0, 0);
            if old {
                php_baseline::emit_php_pack(&mut chunks, 0, 2, 0);
            } else {
                vybe_language_php::emitter::packing_adapter::emit_php_pack(&mut chunks, 0, 2, 0);
            }
            chunks[0].emit_op(Op::RETURN, 0);
            bundle::finalize_with_runtime_helpers(&mut chunks);
            globals::declare_free_globals(&mut chunks);
            globals::normalize_global_table(&mut chunks);
            chunks
        };
        let old = build(true);
        let new = build(false);
        validate(&new);
        assert_eq!(
            vm(0.0).run(old).unwrap().as_str(),
            vm(0.0).run(new).unwrap().as_str(),
            "{format:?}"
        );
    }
}

fn php_unpack_program(old: bool, format: &Value, input: &str, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<php-unpack>")];
    for call in 0..calls {
        vybe_compiler::primitives::instructions::core_wasm::emit_value(&mut chunks[0], 0, format);
        chunks[0].emit_string_const(input, 0);
        if old {
            php_baseline::emit_php_unpack(&mut chunks, 0, 2, 0);
        } else {
            vybe_language_php::emitter::packing_adapter::emit_php_unpack(&mut chunks, 0, 2, 0);
        }
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

fn map_entries(value: Value) -> Vec<(Value, Value)> {
    let Value::Object(object) = value else {
        panic!("expected map")
    };
    let object = object.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Map(entries) = &object.kind else {
        panic!("expected map")
    };
    entries
        .iter()
        .map(|(key, value)| (key.clone(), value.clone()))
        .collect()
}

#[test]
fn shared_php_unpack_preserves_endian_values_named_keys_and_dynamic_formats() {
    for (format, input, expected) in [
        (
            "N",
            "\u{ff}\u{80}\0\u{1}",
            vec![(Value::F64(1.0), Value::F64(4286578689.0))],
        ),
        (
            "V",
            "\u{ff}\u{80}\0\u{1}",
            vec![(Value::F64(1.0), Value::F64(16810239.0))],
        ),
        (
            "n",
            "\u{ff}\u{80}",
            vec![(Value::F64(1.0), Value::F64(65408.0))],
        ),
        (
            "v",
            "\u{ff}\u{80}",
            vec![(Value::F64(1.0), Value::F64(33023.0))],
        ),
        (
            "Cfirst/Csecond",
            "\0\u{ff}",
            vec![
                (Value::String("first".into()), Value::I32(0)),
                (Value::String("second".into()), Value::I32(255)),
            ],
        ),
        (
            "C*",
            "\0\u{ff}",
            vec![
                (Value::I32(1), Value::I32(0)),
                (Value::I32(2), Value::I32(255)),
            ],
        ),
        ("C*", "", vec![]),
    ] {
        for old in [true, false] {
            let chunks = php_unpack_program(old, &Value::String(format.into()), input, 2);
            validate(&chunks);
            assert_eq!(
                map_entries(vm(0.0).run(chunks).unwrap()),
                expected,
                "{format}/{old}"
            );
        }
    }
    for format in [
        Value::Null,
        Value::Bool(false),
        Value::Bool(true),
        Value::F64(0.0),
        Value::F64(1.0),
    ] {
        let old = php_unpack_program(true, &format, "ABCD", 1);
        let new = php_unpack_program(false, &format, "ABCD", 1);
        validate(&new);
        assert_eq!(
            map_entries(vm(0.0).run(new).unwrap()),
            map_entries(vm(0.0).run(old).unwrap())
        );
    }
}

#[test]
fn repeated_php_unpack_shares_body_reduces_module_size_and_finalizes_idempotently() {
    let format = Value::String("N".into());
    let old = php_unpack_program(true, &format, "\u{12}4Vx", 20);
    let mut new = php_unpack_program(false, &format, "\u{12}4Vx", 20);
    validate(&new);
    let before: usize = old.iter().map(|c| c.code.len()).sum();
    let after: usize = new.iter().map(|c| c.code.len()).sum();
    let old_wasm = vybe_platform_wasm::write_wasm(&old);
    let new_wasm = vybe_platform_wasm::write_wasm(&new);
    println!(
        "PHP unpack 20 calls: bytecode={before}/{after} wasm={}/{}",
        old_wasm.len(),
        new_wasm.len()
    );
    assert!(after < before);
    assert!(new_wasm.len() < old_wasm.len());
    assert_eq!(
        new.iter()
            .filter(|c| c.name == "<vybe:php:unpack:2>")
            .count(),
        1
    );
    bundle::finalize_with_runtime_helpers(&mut new);
    assert_eq!(vybe_platform_wasm::write_wasm(&new), new_wasm);
    assert_eq!(
        map_entries(vm(0.0).run(new).unwrap()),
        map_entries(vm(0.0).run(old).unwrap())
    );
}

#[test]
fn php_unpack_frontend_links_pack_and_unpack_helpers_without_clobbering_arguments() {
    vybe_language_php::register();
    for (format, values, expression, expected) in [
        ("N", "305419896", "$result[1]", "305419896"),
        ("V", "4294967295", "$result[1]", "4294967295"),
        ("n", "65408", "$result[1]", "65408"),
        ("v", "33023", "$result[1]", "33023"),
        (
            "C2",
            "0, 255",
            "$result['first'] . ',' . $result['second']",
            "0,255",
        ),
    ] {
        let unpack_format = if format == "C2" {
            "Cfirst/Csecond"
        } else {
            format
        };
        let source = format!(
            "<?php $packFormat = '{format}'; $unpackFormat = '{unpack_format}'; $data = pack($packFormat, {values}); $result = unpack($unpackFormat, $data); return {expression};"
        );
        let module = vybe_language_php::parse(&source).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        validate(&chunks);
        assert_eq!(
            vm(0.0).run(chunks).unwrap().to_string(),
            expected,
            "{source}"
        );
    }
}

#[test]
#[ignore = "manual interleaved PHP unpack diagnostic benchmark"]
fn php_unpack_execution_benchmark() {
    let format = Value::String("N".into());
    let programs = [
        php_unpack_program(true, &format, "\u{12}4Vx", 128),
        php_unpack_program(false, &format, "\u{12}4Vx", 128),
    ];
    let mut times = [Vec::new(), Vec::new()];
    for repeat in 0..5 {
        for which in if repeat % 2 == 0 { [0, 1] } else { [1, 0] } {
            let mut vm = vm(0.0);
            let chunks = programs[which].clone();
            let start = std::time::Instant::now();
            let result = vm.run(chunks).unwrap();
            times[which].push(start.elapsed().as_micros());
            assert_eq!(
                map_entries(result),
                vec![(Value::F64(1.0), Value::F64(305419896.0))]
            );
        }
    }
    times[0].sort();
    times[1].sort();
    println!(
        "PHP unpack diagnostic: before={}us after={}us speedup={:.2}x",
        times[0][2],
        times[1][2],
        times[0][2] as f64 / times[1][2] as f64
    );
}
