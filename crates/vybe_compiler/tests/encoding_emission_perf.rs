//! Encoding producer specialization, private output assembly, and regressions.
use vybe_compiler::primitives::{bundle, globals, platforms, string_encoding as optimized};
use vybe_runtime::{Chunk, Op, VM, Value};
use vybex as _;

#[path = "fixtures/dynamic_string_encoding.rs"]
mod baseline;

type Emitter = fn(&mut [Chunk], usize, u8, u32);
fn program(emitter: Emitter) -> Vec<Chunk> {
    let mut script = Chunk::new("<codec>");
    globals::emit_read(&mut script, "input", 0);
    emitter(std::slice::from_mut(&mut script), 0, 1, 0);
    script.emit_op(Op::RETURN, 0);
    let mut chunks = vec![script];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn runtime(input: &str) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", Value::String(input.into()));
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}

#[test]
fn boxed_bigint_constants_preserve_digits_and_validate() {
    for text in [
        "9223372036854775807",
        "-9223372036854775808",
        "184467440737095516160000",
    ] {
        let value = Value::bigint(vybe_runtime::bigint::BigIntVal::parse(text).unwrap());
        for central in [false, true] {
            let mut script = Chunk::new("<bigint-constant>");
            if central {
                vybe_compiler::primitives::datetime::push_const(&mut script, value.clone(), 0);
            } else {
                vybe_compiler::primitives::instructions::core_wasm::emit_value(
                    &mut script,
                    0,
                    &value,
                );
            }
            let to_string = script.add_import("ecma:bigint", "toString");
            script.emit_call(to_string, 1, 0);
            script.emit_op(Op::RETURN, 0);
            let chunks = vec![script];
            validate(&chunks);
            assert_eq!(runtime("").run(chunks).unwrap().as_str(), text);
        }
    }
}
fn codecs() -> [(&'static str, Emitter, Emitter, &'static str); 8] {
    [
        (
            "bin2hex",
            baseline::emit_bin2hex,
            optimized::emit_bin2hex,
            "Cat\0=ÿ",
        ),
        (
            "hex2bin",
            baseline::emit_hex2bin,
            optimized::emit_hex2bin,
            "436174003Dff",
        ),
        (
            "rot13",
            baseline::emit_str_rot13,
            optimized::emit_str_rot13,
            "Abz 13! ÿ",
        ),
        (
            "qp-encode",
            baseline::emit_quoted_printable_encode,
            optimized::emit_quoted_printable_encode,
            "one=two!\t\0ÿ",
        ),
        (
            "qp-decode",
            baseline::emit_quoted_printable_decode,
            optimized::emit_quoted_printable_decode,
            "one=3Dtwo=0D=0A",
        ),
        (
            "uu-encode",
            baseline::emit_convert_uuencode,
            optimized::emit_convert_uuencode,
            "Cat",
        ),
        (
            "uu-decode",
            baseline::emit_convert_uudecode,
            optimized::emit_convert_uudecode,
            "#0V%T\n",
        ),
        (
            "base64-lenient",
            baseline::emit_base64_decode,
            optimized::emit_base64_decode,
            "S G!VsbG8=\n",
        ),
    ]
}
#[test]
fn scalar_and_string_specialization_matches_captured_emitters() {
    for (name, before, after, input) in codecs() {
        let old = program(before);
        let new = program(after);
        validate(&old);
        validate(&new);
        for input in ["", input] {
            if name == "uu-decode" && input.is_empty() {
                continue;
            }
            let expected = runtime(input).run(old.clone()).unwrap();
            let actual = runtime(input).run(new.clone()).unwrap();
            assert!(
                matches!((&actual, &expected), (Value::String(_), Value::String(_))),
                "{name} {input:?}: {actual:?} / {expected:?}"
            );
            assert_eq!(actual.as_str(), expected.as_str(), "{name} {input:?}");
        }
        println!(
            "encoding {name}: bytecode={}/{} wasm={}/{}",
            old.iter().map(|c| c.code.len()).sum::<usize>(),
            new.iter().map(|c| c.code.len()).sum::<usize>(),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}
#[test]
fn hex_decoder_rejects_odd_length_and_non_hex_digits() {
    let chunks = program(optimized::emit_hex2bin);
    validate(&chunks);
    for input in ["f", "123", "00:", "GG", "7z", "/0"] {
        assert!(
            matches!(
                runtime(input).run(chunks.clone()).unwrap(),
                Value::Bool(false)
            ),
            "{input}"
        );
    }
    assert_eq!(runtime("417A00").run(chunks).unwrap().as_str(), "Az\0");
}
#[test]
fn quoted_printable_decoder_preserves_invalid_escapes_and_bounds_reads() {
    let chunks = program(optimized::emit_quoted_printable_decode);
    validate(&chunks);
    // These expected results were checked against PHP quoted_printable_decode.
    for (input, expected) in [
        ("=G0", "=G0"),
        ("=0G", "=0G"),
        ("=A", "=A"),
        ("=", ""),
        ("= ", ""),
        ("=\t", ""),
        ("a=\nb", "ab"),
        ("a=\rXb", "aXb"),
        ("a= \nb", "ab"),
        ("a=\t\r\nb", "ab"),
        ("a= \rXb", "aXb"),
        ("a= \tXb", "a= \tXb"),
        ("a=\r", "a"),
        ("a=\n", "a"),
        ("one=3Dtwo=0D=0A", "one=two\r\n"),
        ("=af", "¯"),
    ] {
        assert_eq!(
            runtime(input).run(chunks.clone()).unwrap().as_str(),
            expected,
            "{input:?}"
        );
    }
}
#[test]
#[ignore = "execution benchmark; run with --include-ignored --nocapture"]
fn encoding_execution_benchmark() {
    use std::time::Instant;
    for (name, before, after, _) in codecs()
        .into_iter()
        .take(5)
        .filter(|(name, ..)| *name != "hex2bin")
    {
        let programs = [program(before), program(after)];
        for size in [256, 8192] {
            let input = if name == "qp-decode" {
                "=41x".repeat(size / 4)
            } else {
                "Abz=09!\t".repeat(size / 8)
            };
            let mut samples = [Vec::new(), Vec::new()];
            let mut expected = None;
            for _ in 0..5 {
                for (which, chunks) in programs.iter().enumerate() {
                    let mut vm = runtime(&input);
                    let start = Instant::now();
                    let result = vm.run(chunks.clone()).unwrap();
                    samples[which].push(start.elapsed().as_micros());
                    let output = result.as_str().to_owned();
                    if let Some(expected) = &expected {
                        assert_eq!(&output, expected);
                    } else {
                        expected = Some(output);
                    }
                }
            }
            for sample in &mut samples {
                sample.sort_unstable();
            }
            println!(
                "encoding-execution {name} size={size}: dynamic={}us specialized={}us speedup={:.2}x",
                samples[0][2],
                samples[1][2],
                samples[0][2] as f64 / samples[1][2] as f64
            );
        }
    }
}

#[test]
fn uudecode_bounds_and_six_bit_values_match_php() {
    let chunks = program(optimized::emit_convert_uudecode);
    validate(&chunks);
    for input in ["", "#", "#0V", "!00", "~0V%T\n"] {
        assert!(
            matches!(
                runtime(input).run(chunks.clone()).unwrap(),
                Value::Bool(false)
            ),
            "{input:?}"
        );
    }
    for (input, expected) in [
        ("`", ""),
        ("`\n", ""),
        (" \n", ""),
        ("#0V%T", "Cat"),
        ("#0V%T\n", "Cat"),
        ("!0000", "A"),
        ("a0000", "A"),
        ("!~~~~", "y"),
        ("!zzzz", "i"),
    ] {
        assert_eq!(
            runtime(input).run(chunks.clone()).unwrap().as_str(),
            expected,
            "{input:?}"
        );
    }
}

#[test]
fn failed_decoders_return_to_the_caller_instead_of_returning_from_it() {
    for (emitter, input) in [
        (optimized::emit_hex2bin as Emitter, "123"),
        (optimized::emit_hex2bin as Emitter, "GG"),
        (optimized::emit_convert_uudecode as Emitter, "!00"),
    ] {
        let mut script = Chunk::new("<caller>");
        script.emit_string_const(input, 0);
        emitter(std::slice::from_mut(&mut script), 0, 1, 0);
        script.emit_op(Op::DROP, 0);
        script.emit_string_const("continued", 0);
        script.emit_op(Op::RETURN, 0);
        let mut chunks = vec![script];
        bundle::finalize_with_runtime_helpers(&mut chunks);
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        validate(&chunks);
        assert_eq!(runtime("").run(chunks).unwrap().as_str(), "continued");
    }
}

#[test]
fn explicit_base64_strict_flag_retains_dynamic_truthiness() {
    let mut script = Chunk::new("<strict-codec>");
    globals::emit_read(&mut script, "input", 0);
    globals::emit_read(&mut script, "strict", 0);
    optimized::emit_base64_decode(std::slice::from_mut(&mut script), 0, 2, 0);
    script.emit_op(Op::RETURN, 0);
    let mut chunks = vec![script];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate(&chunks);
    for (flag, strict) in [
        (Value::Bool(false), false),
        (Value::Bool(true), true),
        (Value::I32(0), false),
        (Value::I32(1), true),
        (Value::String("".into()), false),
        (Value::String("yes".into()), true),
    ] {
        let mut vm = runtime("!!!!");
        vm.set_global("strict", flag);
        let result = vm.run(chunks.clone()).unwrap();
        if strict {
            assert!(matches!(result, Value::Bool(false)));
        } else {
            assert!(matches!(result, Value::String(_)));
            assert_eq!(result.as_str(), "");
        }
    }
}

#[test]
fn shared_codecs_preserve_repeated_calls_and_finalization() {
    for (name, before, after, input) in codecs() {
        let mut sizes = Vec::new();
        for emitter in [before, after] {
            let mut script = Chunk::new("<repeated-codec>");
            for _ in 0..20 {
                globals::emit_read(&mut script, "input", 0);
                emitter(std::slice::from_mut(&mut script), 0, 1, 0);
                script.emit_op(Op::DROP, 0);
            }
            script.emit_string_const("done", 0);
            script.emit_op(Op::RETURN, 0);
            let mut chunks = vec![script];
            bundle::finalize_with_runtime_helpers(&mut chunks);
            let count = chunks.len();
            let code: Vec<_> = chunks.iter().map(|c| c.code.clone()).collect();
            bundle::finalize_with_runtime_helpers(&mut chunks);
            assert_eq!(chunks.len(), count);
            assert_eq!(
                chunks.iter().map(|c| c.code.clone()).collect::<Vec<_>>(),
                code
            );
            globals::declare_free_globals(&mut chunks);
            globals::normalize_global_table(&mut chunks);
            validate(&chunks);
            sizes.push((
                chunks.iter().map(|c| c.code.len()).sum::<usize>(),
                vybe_platform_wasm::write_wasm(&chunks).len(),
            ));
            assert_eq!(runtime(input).run(chunks).unwrap().as_str(), "done");
        }
        assert!(sizes[1].0 * 4 < sizes[0].0, "{name}: {sizes:?}");
        assert!(sizes[1].1 * 4 < sizes[0].1, "{name}: {sizes:?}");
        println!(
            "encoding-repeated {name} calls=20: bytecode={}/{} wasm={}/{}",
            sizes[0].0, sizes[1].0, sizes[0].1, sizes[1].1
        );
    }
}

#[test]
fn php_codec_frontend_preserves_results_and_failure_continuations() {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    for (body, expected) in [
        ("return bin2hex('Cat');", "436174"),
        ("return bin2hex(true);", "31"),
        ("return bin2hex(false);", ""),
        ("return bin2hex(null);", ""),
        ("return bin2hex(42);", "3432"),
        ("return bin2hex(1.5);", "312e35"),
        (
            "return bin2hex(9223372036854775807);",
            "39323233333732303336383534373735383037",
        ),
        ("return hex2bin('436174');", "Cat"),
        (
            "return hex2bin('123') === false ? 'continued' : 'wrong';",
            "continued",
        ),
        (
            "return hex2bin('GG') === false ? 'continued' : 'wrong';",
            "continued",
        ),
        (
            "return convert_uudecode('!00') === false ? 'continued' : 'wrong';",
            "continued",
        ),
        ("return quoted_printable_decode('=G0');", "=G0"),
        (r#"return quoted_printable_decode("a=\rXb");"#, "aXb"),
        ("return str_rot13('Abz');", "Nom"),
        ("return convert_uudecode('#0V%T');", "Cat"),
        ("return base64_decode('S G!VsbG8=');", "Hello"),
    ] {
        let module = vybe_language_php::parse(&format!("<?php {body}")).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        let binary = vybe_platform_wasm::write_wasm(&chunks);
        if let Err(error) =
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&binary)
        {
            std::fs::write(
                std::env::temp_dir().join("encoding-frontend-invalid.wasm"),
                &binary,
            )
            .unwrap();
            panic!("{body}: {error}");
        }
        assert_eq!(
            runtime("").run(chunks).unwrap().as_str(),
            expected,
            "{body}"
        );
    }
}

#[test]
#[ignore = "manual export for independent wasm-tools validation"]
fn export_encoding_wasm_fixtures() {
    let dir = std::env::temp_dir().join("vybe-encoding-wasm");
    std::fs::create_dir_all(&dir).unwrap();
    for (name, _, emitter, _) in codecs() {
        let chunks = program(emitter);
        std::fs::write(
            dir.join(format!("{name}.wasm")),
            vybe_platform_wasm::write_wasm(&chunks),
        )
        .unwrap();
    }
}
