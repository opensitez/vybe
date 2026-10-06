use vybe_compiler::primitives::{bundle, globals, platforms, string_encoding};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/text_optimization_base64_predicate_selected.rs"]
mod candidate;
#[path = "fixtures/text_optimization_base64_predicate_frozen.rs"]
mod frozen;
const LENIENT: &str = "__stdlib_encoding_base64_decode_lenient";
const STRICT: &str = "__stdlib_encoding_base64_decode";

fn ops(chunk: &Chunk) -> Vec<(usize, Op)> {
    let mut offset = 0;
    let mut result = vec![];
    while offset < chunk.code.len() {
        let op = Op::from_code_at(&chunk.code, offset);
        result.push((offset, op));
        offset += 4 + op.operand_format().size_in(&chunk.code, offset + 4);
    }
    result
}
fn u16_at(c: &Chunk, at: usize) -> u16 {
    u16::from_be_bytes(c.code[at + 4..at + 6].try_into().unwrap())
}
fn validate(chunks: &[Chunk]) -> usize {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    bytes.len()
}
fn vm() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}
fn run(vm: &mut VM, chunks: Vec<Chunk>) -> Result<Value, String> {
    vm.run(chunks).map_err(|mut e| {
        for f in &mut e.call_stack {
            f.offset = 0;
        }
        e.to_string()
    })
}
fn select_candidate(original: &[Chunk]) -> Vec<Chunk> {
    let mut result = original.to_vec();
    for index in 0..original.len() {
        if [LENIENT, STRICT].contains(&original[index].name.as_str()) {
            result = select_one(&result, index, false);
        }
    }
    result
}
fn select_one(original: &[Chunk], index: usize, inverse: bool) -> Vec<Chunk> {
    let source = &original[index];
    let decoded = ops(source);
    let prefix = source.arity as usize;
    assert!([1, 2].contains(&prefix));
    for (slot, &(at, op)) in decoded[..prefix].iter().enumerate() {
        assert_eq!(op, Op::LOCAL_GET);
        assert_eq!(u16_at(source, at), slot as u16);
    }
    let offset = decoded[prefix].0;
    let base = source.arity as u16;
    let line = source.lines[offset];
    assert!(!source.is_async && !source.is_generator);
    assert!(source.block_i32_results.iter().all(|&k| k < offset));
    assert!(source.call_indirect_sigs.keys().all(|&k| k < offset));
    assert!(source.call_indirect_canon.keys().all(|&k| k < offset));
    assert!(source.stack_switch_handlers.keys().all(|&k| k < offset));
    assert!(
        source
            .user_code_offset
            .is_none_or(|k| (k as usize) < offset)
    );
    let prepare = || {
        let mut c = source.clone();
        c.code.truncate(offset);
        c.lines.truncate(offset);
        c.local_count = base;
        c
    };
    let mut result = original.to_vec();
    let relocate = |c: &mut Chunk| {
        assert_eq!(c.constants, source.constants, "no constant-pool rebinding");
        assert_eq!(c.global_imports, source.global_imports);
        for (at, op) in ops(c).into_iter().filter(|(at, _)| *at >= offset) {
            if op == Op::CALL {
                let import = c.imports[u16_at(c, at) as usize].clone();
                // Keep the actual module import carrier. An unlinked helper
                // owns its local table; a linked helper has an empty table and
                // indexes the root. Rebuilding must not promote a second copy.
                let carrier = if source.imports.is_empty() {
                    &original[0].imports
                } else {
                    &source.imports
                };
                let idx = carrier
                    .iter()
                    .position(|i| *i == import)
                    .expect("predicate replacement introduces no new host imports");
                c.code[at + 4..at + 6].copy_from_slice(&u16::try_from(idx).unwrap().to_be_bytes());
            } else if op == Op::GLOBAL_GET {
                let idx = u32::from_be_bytes(c.code[at + 4..at + 8].try_into().unwrap()) as usize;
                let Value::String(name) = &c.constants[idx] else {
                    panic!("expected string-constant key")
                };
                let idx = source
                    .globals
                    .iter()
                    .position(|key| key == name.as_ref())
                    .unwrap();
                c.code[at + 4..at + 8].copy_from_slice(&(idx as u32).to_be_bytes());
            } else {
                assert_ne!(op, Op::GLOBAL_SET, "helper has no global writes");
            }
        }
        c.imports = source.imports.clone();
    };
    let mut baseline = prepare();
    if inverse {
        candidate::emit_body(std::slice::from_mut(&mut baseline), 0, source.arity, line);
    } else {
        frozen::emit_body(std::slice::from_mut(&mut baseline), 0, source.arity, line);
    }
    relocate(&mut baseline);
    let end = baseline.code.len();
    assert_eq!(
        baseline.code[offset..],
        source.code[offset..end],
        "exact actual frozen body"
    );
    assert_eq!(
        baseline.local_count,
        base + if source.arity == 1 { 6 } else { 12 }
    );
    assert_eq!(
        decoded
            .iter()
            .filter(|(at, _)| *at >= end)
            .map(|(_, op)| *op)
            .collect::<Vec<_>>(),
        vec![Op::RETURN]
    );
    let mut changed = prepare();
    if inverse {
        frozen::emit_body(std::slice::from_mut(&mut changed), 0, source.arity, line);
    } else {
        candidate::emit_body(std::slice::from_mut(&mut changed), 0, source.arity, line);
    }
    relocate(&mut changed);
    assert_eq!(
        changed.local_count, baseline.local_count,
        "same default/strict scratch budget"
    );
    assert_eq!(changed.scratch_high_water, source.scratch_high_water);
    let new_end = changed.code.len();
    changed.code.extend_from_slice(&source.code[end..]);
    changed.lines.extend_from_slice(&source.lines[end..]);
    assert_eq!(&changed.code[..offset], &source.code[..offset]);
    assert_eq!(&changed.lines[..offset], &source.lines[..offset]);
    assert_eq!(&changed.code[new_end..], &source.code[end..]);
    assert_eq!(changed.lines.len(), changed.code.len());
    assert_eq!(changed.arity, source.arity);
    assert_eq!(changed.param_count, source.param_count);
    assert_eq!(changed.takes_receiver, source.takes_receiver);
    assert_eq!(changed.module_receiver_abi, source.module_receiver_abi);
    assert_eq!(changed.capture_count, source.capture_count);
    assert_eq!(changed.block_i32_results, source.block_i32_results);
    assert_eq!(changed.call_indirect_sigs, source.call_indirect_sigs);
    assert_eq!(changed.call_indirect_canon, source.call_indirect_canon);
    assert_eq!(changed.stack_switch_handlers, source.stack_switch_handlers);
    result[index] = changed;
    result
}
fn frozen_program(production: &[Chunk]) -> Vec<Chunk> {
    let mut result = production.to_vec();
    for index in 0..production.len() {
        if [LENIENT, STRICT].contains(&production[index].name.as_str()) {
            result = select_one(&result, index, true);
        }
    }
    result
}
fn program(calls: usize, strict: bool, retained: bool, shadow: bool) -> Vec<Chunk> {
    frozen_program(&production_program(calls, strict, retained, shadow))
}
fn parsed_php(body: &str) -> Vec<Chunk> {
    frozen_program(&production_parsed_php(body))
}
fn production_program(calls: usize, strict: bool, retained: bool, shadow: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<base64-ascii-predicate>")];
    let binding = if strict {
        "__vybe_encoding_base64_decode"
    } else {
        "__vybe_encoding_base64_decode_lenient"
    };
    if shadow {
        globals::emit_read(&mut chunks[0], "replacement", 0);
        globals::emit_write(&mut chunks[0], binding, 0);
    }
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "input", 0);
        if strict {
            globals::emit_read(&mut chunks[0], "strict", 0);
        }
        string_encoding::emit_base64_decode(&mut chunks, 0, if strict { 2 } else { 1 }, 0);
        if retained || call + 1 < calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    if retained {
        globals::emit_read(&mut chunks[0], binding, 0);
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn production_parsed_php(body: &str) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    let module = vybe_language_php::parse(&format!("<?php {body}")).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}
#[test]
fn installed_production_matches_selected_emission_and_source_gate() {
    for strict in [false, true] {
        let production = production_program(1, strict, false, false);
        let frozen = frozen_program(&production);
        let rebuilt = select_candidate(&frozen);
        validate(&production);
        validate(&frozen);
        validate(&rebuilt);
        assert_eq!(
            production.iter().map(|c| &c.code).collect::<Vec<_>>(),
            rebuilt.iter().map(|c| &c.code).collect::<Vec<_>>()
        );
        assert_eq!(production[0].imports, rebuilt[0].imports);
        assert_eq!(production[0].global_imports, rebuilt[0].global_imports);
        for (actual, reconstructed) in production.iter().zip(&rebuilt) {
            assert_eq!(actual.imports, reconstructed.imports);
            assert_eq!(actual.global_imports, reconstructed.global_imports);
            assert_eq!(actual.constants, reconstructed.constants);
            assert_eq!(actual.lines, reconstructed.lines);
        }
        assert_eq!(
            vybe_platform_wasm::write_wasm(&production),
            vybe_platform_wasm::write_wasm(&rebuilt),
            "complete standard module equality, including resolved CALL imports"
        );
        let options = vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: true,
            include_name_section: true,
        };
        assert_eq!(
            vybe_platform_wasm::write_wasm_with_options(&production, options),
            vybe_platform_wasm::write_wasm_with_options(&rebuilt, options),
            "complete module plus serialized bytecode/debug/handler metadata equality"
        );
        let input = Value::String("QQ$==".into());
        for chunks in [production, rebuilt] {
            let mut runtime = vm();
            runtime.set_global("input", input.clone());
            runtime.set_global("strict", Value::Bool(true));
            assert_eq!(
                run(&mut runtime, chunks),
                Ok(if strict {
                    Value::Bool(false)
                } else {
                    Value::String("A".into())
                })
            );
        }
    }
}
#[test]
fn all_bytes_unicode_flags_padding_errors_and_native() {
    for strict in [false, true] {
        let old = program(1, strict, false, false);
        let new = select_candidate(&old);
        let name = if strict { STRICT } else { LENIENT };
        let a = old.iter().find(|c| c.name == name).unwrap();
        let b = new.iter().find(|c| c.name == name).unwrap();
        println!(
            "base64 {name} body={}→{} ops={}→{} native={}→{} locals={}→{}",
            a.code.len(),
            b.code.len(),
            ops(a).len(),
            ops(b).len(),
            validate(&old),
            validate(&new),
            a.local_count,
            b.local_count
        );
        assert!(b.code.len() < a.code.len());
        assert!(ops(b).len() < ops(a).len());
        let mut before = vm();
        let mut after = vm();
        let mut cases = (0..=255)
            .map(|c| format!("Q{}Q==", char::from_u32(c).unwrap()))
            .collect::<Vec<_>>();
        cases.extend(
            [
                "", "QQ==", "QUI=", "QUJD", "Q Q\n==", "QQ", "Q", "====", "QQ===", "Q=Q=",
                "QQ==junk", "é", "😀", "QQ😀==", "ĀQQ==", "🦀",
            ]
            .map(str::to_owned),
        );
        cases.extend(["QUJD".repeat(2048), "Q Q\n".repeat(64)]);
        let flags = if strict {
            vec![
                Value::Bool(true),
                Value::Bool(false),
                Value::Null,
                Value::I32(0),
                Value::I32(1),
                Value::F64(f64::NAN),
                Value::String("".into()),
                Value::String("false".into()),
            ]
        } else {
            vec![Value::Bool(true)]
        };
        let small = vec![
            "QQ==".to_owned(),
            "QQ$==".to_owned(),
            "QQ😀==".to_owned(),
            "".to_owned(),
        ];
        for flag in flags {
            let inputs = if matches!(flag, Value::Bool(_)) {
                &cases
            } else {
                &small
            };
            for text in inputs {
                for runtime in [&mut before, &mut after] {
                    runtime.set_global("input", Value::String(text.clone().into()));
                    runtime.set_global("strict", flag.clone());
                }
                assert_eq!(
                    run(&mut before, old.clone()),
                    run(&mut after, new.clone()),
                    "strictarity={strict} flag={flag:?}, text={text:?}"
                );
            }
        }
        for value in [
            Value::Null,
            Value::Undefined,
            Value::I32(42),
            Value::Bool(true),
        ] {
            for runtime in [&mut before, &mut after] {
                runtime.set_global("input", value.clone());
                runtime.set_global("strict", Value::Bool(true));
            }
            let a = run(&mut before, old.clone());
            let b = run(&mut after, new.clone());
            assert!(a.is_err());
            assert_eq!(a, b);
        }
        // Independent language expectations start with fresh execution state.
        // The reused-VM differential error/flag matrix above remains intact.
        for (text, flag) in [("QQ==", true), ("QQ$==", true), ("QQ$==", false)] {
            let expected = if strict && flag && text.contains('$') {
                Value::Bool(false)
            } else {
                Value::String("A".into())
            };
            for (label, chunks) in [("baseline", &old), ("predicate", &new)] {
                let mut runtime = vm();
                runtime.set_global("input", Value::String(text.into()));
                runtime.set_global("strict", Value::Bool(flag));
                assert_eq!(
                    run(&mut runtime, chunks.clone()),
                    Ok(expected.clone()),
                    "standalone {label}: strictABI={strict}, flag={flag}, input={text:?}"
                );
            }
        }
    }
}
#[test]
fn actual_php_scalar_flags_argument_trace_and_public_override() {
    for (body, expected) in [
        ("return base64_decode('QQ==');", Value::String("A".into())),
        ("return base64_decode('QQ$==',true);", Value::Bool(false)),
        (
            "return base64_decode('QQ$==',false);",
            Value::String("A".into()),
        ),
        ("return base64_decode(null);", Value::String("".into())),
        (
            "$order=''; function text() { global $order; $order.='t'; return 'QQ=='; } function flag() { global $order; $order.='f'; return true; } $r=base64_decode(text(),flag()); return $order.'|'.$r;",
            Value::String("tf|A".into()),
        ),
    ] {
        let old = parsed_php(body);
        let new = select_candidate(&old);
        validate(&old);
        validate(&new);
        assert_eq!(run(&mut vm(), old), Ok(expected.clone()), "old {body}");
        assert_eq!(run(&mut vm(), new), Ok(expected), "new {body}");
    }
    use std::sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    };
    for strict in [false, true] {
        for shadow in [false, true] {
            let old = program(1, strict, false, shadow);
            let new = select_candidate(&old);
            validate(&old);
            validate(&new);
            for chunks in [old, new] {
                let mut runtime = vm();
                let seen = Arc::new(AtomicUsize::new(0));
                let count = seen.clone();
                runtime.register_host_fn(
                    "test:b64",
                    "replace",
                    Box::new(move |_, _| {
                        count.fetch_add(1, Ordering::SeqCst);
                        Value::String("replacement".into())
                    }),
                );
                let idx = runtime
                    .resolve_host_function_index("test:b64", "replace")
                    .unwrap();
                let mut function = vybe_runtime::value::Object::new();
                function.kind = vybe_runtime::value::ObjectKind::HostFunction(idx);
                let callable = Value::Object(vybe_runtime::heap::alloc(function));
                let mut object = vybe_runtime::value::Object::new();
                object
                    .properties
                    .insert("Symbol(toPrimitive)".into(), callable.clone());
                runtime.set_global("input", Value::Object(vybe_runtime::heap::alloc(object)));
                runtime.set_global("strict", Value::Bool(true));
                runtime.set_global("replacement", callable);
                let observed = run(&mut runtime, chunks);
                if shadow {
                    assert_eq!(observed, Ok(Value::String("replacement".into())));
                    assert_eq!(seen.load(Ordering::SeqCst), 1);
                } else {
                    assert!(observed.unwrap_err().contains("not a string"));
                    assert_eq!(seen.load(Ordering::SeqCst), 0);
                }
            }
        }
    }
}
#[test]
fn scan_all_units_even_after_first_invalid_and_repeated_calls() {
    use std::sync::{Arc, Mutex};
    for strict in [false, true] {
        let old = program(16, strict, false, false);
        let new = select_candidate(&old);
        validate(&old);
        validate(&new);
        assert_eq!(old[0].code, new[0].code);
        let name = if strict { STRICT } else { LENIENT };
        assert_eq!(new.iter().filter(|c| c.name == name).count(), 1);
        for calls in [1, 16] {
            let old = program(calls, strict, false, false);
            let new = select_candidate(&old);
            for chunks in [old, new] {
                let mut runtime = vm();
                let seen = Arc::new(Mutex::new(vec![]));
                let log = seen.clone();
                runtime.register_host_fn(
                    "wasm:js-string",
                    "charCodeAt",
                    Box::new(move |_, args| {
                        let Value::String(text) = &args[0] else {
                            panic!("String");
                        };
                        let i = args[1].as_i32();
                        log.lock().unwrap().push(i);
                        Value::I32(text.encode_utf16().nth(i as usize).unwrap() as i32)
                    }),
                );
                runtime.set_global("input", Value::String("$QQ==".into()));
                runtime.set_global("strict", Value::Bool(true));
                let result = run(&mut runtime, chunks);
                if strict {
                    assert_eq!(result, Ok(Value::Bool(false)));
                } else {
                    assert_eq!(result, Ok(Value::String("A".into())));
                }
                assert_eq!(
                    *seen.lock().unwrap(),
                    (0..calls).flat_map(|_| 0..5).collect::<Vec<_>>()
                );
            }
        }
    }
}
#[test]
fn retained_decoded_output_identity_policy_unchanged() {
    for strict in [false, true] {
        let old = program(1, strict, true, false);
        let new = select_candidate(&old);
        validate(&old);
        validate(&new);
        for text in ["", "QQ==", "MA==", "QUI="] {
            let mut identities = vec![];
            for chunks in [&old, &new] {
                let mut runtime = vm();
                runtime.set_global("input", Value::String(text.into()));
                runtime.set_global("strict", Value::Bool(true));
                let function = runtime.run(chunks.clone()).unwrap();
                let args = if strict {
                    vec![Value::String(text.into()), Value::Bool(true)]
                } else {
                    vec![Value::String(text.into())]
                };
                let (Value::String(a), Value::String(b)) = (
                    runtime.invoke(&function, &args).unwrap(),
                    runtime.invoke(&function, &args).unwrap(),
                ) else {
                    panic!("String outputs");
                };
                identities.push(std::sync::Arc::ptr_eq(&a, &b));
            }
            assert_eq!(
                identities[0], identities[1],
                "strict={strict} input={text:?}"
            );
        }
    }
}
fn report_batches(label: &str, mut measure: impl FnMut(usize, usize) -> u128) {
    let mut fastest = u128::MAX;
    for variant in 0..2 {
        for _ in 0..3 {
            fastest = fastest.min(measure(variant, 1));
        }
    }
    let iterations = (30_000_000u128.div_ceil(fastest.max(1)) as usize).clamp(1, 1000);
    let mut raw = [vec![], vec![]];
    for sample in 0..7 {
        for variant in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
            raw[variant].push(measure(variant, iterations));
        }
    }
    let mut paired = raw[0]
        .iter()
        .zip(&raw[1])
        .map(|(a, b)| *a as f64 / *b as f64)
        .collect::<Vec<_>>();
    paired.sort_by(f64::total_cmp);
    println!(
        "base64 predicate {label} iterations={iterations} raw_old_ns={:?} raw_new_ns={:?}",
        raw[0], raw[1]
    );
    for times in &mut raw {
        times.sort();
    }
    println!(
        "base64 predicate {label} paired={:.2}x medians={:.2}x old={}ns new={}ns",
        paired[3],
        raw[0][3] as f64 / raw[1][3] as f64,
        raw[0][3],
        raw[1][3]
    );
}
#[test]
#[ignore = "paired opt0 actual base64 caller/helper cold and retained timings"]
fn benchmark_base64_ascii_predicate() {
    for strict in [false, true] {
        for (name, text, calls) in [
            ("short4", "QQ==".into(), 1),
            ("short64", "QUJD".repeat(16), 1),
            ("1024", "QUJD".repeat(256), 1),
            ("32768", "QUJD".repeat(8192), 1),
            ("junk1024", format!("{}$", "QUJD".repeat(255)), 1),
            ("repeat16", "QUJD".repeat(16), 16),
        ] {
            let old = program(calls, strict, false, false);
            let new = select_candidate(&old);
            validate(&old);
            validate(&new);
            let mut before = vm();
            let mut after = vm();
            let input = Value::String(text.into());
            for runtime in [&mut before, &mut after] {
                runtime.set_global("input", input.clone());
                runtime.set_global("strict", Value::Bool(true));
            }
            let expected = run(&mut before, old.clone());
            let label = format!("{name} strict={strict}");
            report_batches(&format!("{label} cold"), |variant, iterations| {
                let (runtime, chunks) = if variant == 0 {
                    (&mut before, &old)
                } else {
                    (&mut after, &new)
                };
                let mut total = 0;
                for _ in 0..iterations {
                    let prepared = chunks.clone();
                    let start = std::time::Instant::now();
                    let result = run(runtime, prepared);
                    total += start.elapsed().as_nanos();
                    assert_eq!(result, expected);
                }
                total
            });
            let old = program(1, strict, true, false);
            let new = select_candidate(&old);
            let a = before.run(old).unwrap();
            let b = after.run(new).unwrap();
            let args = if strict {
                vec![input, Value::Bool(true)]
            } else {
                vec![input]
            };
            let expected = before.invoke(&a, &args).map_err(|error| error.to_string());
            report_batches(&format!("{label} loaded"), |variant, iterations| {
                let (runtime, function) = if variant == 0 {
                    (&mut before, &a)
                } else {
                    (&mut after, &b)
                };
                let mut outputs = Vec::with_capacity(iterations * calls);
                let start = std::time::Instant::now();
                for _ in 0..iterations {
                    for _ in 0..calls {
                        outputs.push(runtime.invoke(function, &args));
                    }
                }
                let elapsed = start.elapsed().as_nanos();
                for output in outputs {
                    assert_eq!(output.map_err(|error| error.to_string()), expected);
                }
                elapsed
            });
        }
    }
}

fn validate_wasm_bytes(bytes: &[u8]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(bytes)
        .unwrap();
}
fn leb(mut n: usize) -> Vec<u8> {
    let mut out = vec![];
    loop {
        let byte = (n & 127) as u8;
        n >>= 7;
        out.push(if n == 0 { byte } else { byte | 128 });
        if n == 0 {
            return out;
        }
    }
}
fn read_leb(bytes: &[u8], pos: &mut usize) -> usize {
    let (mut n, mut shift) = (0, 0);
    loop {
        let b = bytes[*pos];
        *pos += 1;
        n |= usize::from(b & 127) << shift;
        if b & 128 == 0 {
            return n;
        }
        shift += 7;
        assert!(shift < usize::BITS);
    }
}
fn sections(bytes: &[u8]) -> Vec<(u8, std::ops::Range<usize>, usize)> {
    let mut pos = 8;
    let mut rows = vec![];
    while pos < bytes.len() {
        let start = pos;
        let id = bytes[pos];
        pos += 1;
        let size = read_leb(bytes, &mut pos);
        let payload = pos;
        pos += size;
        assert!(pos <= bytes.len());
        rows.push((id, start..pos, payload));
    }
    rows
}
fn native_export(chunks: &[Chunk], callee: usize, export: &str) -> Vec<u8> {
    assert!(callee > 0 && callee < chunks.len());
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    validate_wasm_bytes(&bytes);
    let (mut imports, mut defined, mut bodies) = (0usize, vec![], 0);
    let mut signatures = vec![];
    let mut exports = vec![];
    for p in wasmparser::Parser::new(0).parse_all(&bytes) {
        match p.unwrap() {
            wasmparser::Payload::TypeSection(section) => {
                for group in section {
                    for ty in group.unwrap().into_types() {
                        signatures.push(match ty.composite_type.inner {
                            wasmparser::CompositeInnerType::Func(f) => {
                                Some((f.params().to_vec(), f.results().to_vec()))
                            }
                            _ => None,
                        });
                    }
                }
            }
            wasmparser::Payload::ImportSection(section) => {
                for i in section.into_imports() {
                    if matches!(
                        i.unwrap().ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                    ) {
                        imports += 1;
                    }
                }
            }
            wasmparser::Payload::FunctionSection(section) => {
                defined.extend(section.into_iter().map(Result::unwrap))
            }
            wasmparser::Payload::CodeSectionEntry(_) => bodies += 1,
            wasmparser::Payload::ExportSection(section) => {
                exports.extend(section.into_iter().map(Result::unwrap))
            }
            _ => (),
        }
    }
    assert_eq!(defined.len(), chunks.len());
    assert_eq!(bodies, chunks.len());
    for adjacent in defined.windows(2) {
        assert_eq!(adjacent[1], adjacent[0] + 1);
    }
    let (params, results) = signatures[defined[callee] as usize].as_ref().unwrap();
    assert_eq!(params.len(), chunks[callee].arity as usize);
    assert!(params.iter().all(|p| *p == wasmparser::ValType::EXTERNREF));
    assert_eq!(results, &[wasmparser::ValType::EXTERNREF]);
    assert!(!exports.iter().any(|e| e.name == export));
    assert!(
        exports
            .iter()
            .any(|e| e.name == "_start" && e.kind == wasmparser::ExternalKind::Func)
    );
    let before = sections(&bytes);
    let mut out = bytes[..8].to_vec();
    for (id, range, payload) in &before {
        if *id != 7 {
            out.extend_from_slice(&bytes[range.clone()]);
            continue;
        }
        let mut pos = *payload;
        let count = read_leb(&bytes, &mut pos);
        assert_eq!(count, exports.len());
        let mut body = leb(count + 1);
        body.extend_from_slice(&bytes[pos..range.end]);
        body.extend(leb(export.len()));
        body.extend(export.as_bytes());
        body.push(0);
        body.extend(leb(imports + callee));
        out.push(7);
        out.extend(leb(body.len()));
        out.extend(body);
    }
    let after = sections(&out);
    assert_eq!(before.len(), after.len());
    for ((id, a, _), (other_id, b, _)) in before.iter().zip(after.iter()) {
        assert_eq!(id, other_id);
        if *id != 7 {
            assert_eq!(&bytes[a.clone()], &out[b.clone()]);
        }
    }
    validate_wasm_bytes(&out);
    out
}
#[test]
#[ignore = "writes genuine full-module exported helpers for separate V8 execution"]
fn export_base64_predicate_native_kernels() {
    let dir = std::path::Path::new(file!())
        .parent()
        .unwrap()
        .join("base64-predicate-native-artifacts");
    std::fs::create_dir_all(&dir).unwrap();
    for strict in [false, true] {
        let old = program(1, strict, false, false);
        let new = select_candidate(&old);
        let name = if strict { STRICT } else { LENIENT };
        for (label, chunks) in [("frozen", old), ("candidate", new)] {
            let index = chunks.iter().position(|c| c.name == name).unwrap();
            let bytes = native_export(&chunks, index, "decode");
            let filename = format!(
                "base64-{}-{label}.wasm",
                if strict { "strict" } else { "default" }
            );
            std::fs::write(dir.join(&filename), bytes).unwrap();
            println!(
                "native export {filename}: actual complete module, helperarity={}",
                chunks[index].arity
            );
        }
    }
}
