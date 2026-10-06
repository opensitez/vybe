use vybe_compiler::primitives::{bundle, globals, platforms};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/text_optimization_htmlentities_selected.rs"]
mod candidate;
#[path = "fixtures/text_optimization_htmlentities_frozen_named.rs"]
mod frozen;

fn ops(c: &Chunk) -> Vec<(usize, Op)> {
    let mut at = 0;
    let mut result = vec![];
    while at < c.code.len() {
        let op = Op::from_code_at(&c.code, at);
        result.push((at, op));
        at += 4 + op.operand_format().size_in(&c.code, at + 4);
    }
    result
}
fn u16_at(c: &Chunk, at: usize) -> usize {
    u16::from_be_bytes(c.code[at + 4..at + 6].try_into().unwrap()) as usize
}
fn u32_at(c: &Chunk, at: usize) -> usize {
    u32::from_be_bytes(c.code[at + 4..at + 8].try_into().unwrap()) as usize
}
fn vm() -> VM {
    let mut runtime = VM::new();
    platforms::register_platforms_all(&mut runtime);
    runtime
}
fn run(runtime: &mut VM, chunks: Vec<Chunk>) -> Result<Value, String> {
    runtime.run(chunks).map_err(|mut error| {
        for frame in &mut error.call_stack {
            frame.offset = 0;
        }
        error.to_string()
    })
}
fn validate(chunks: &[Chunk]) -> usize {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        let selected = chunks.iter().any(|chunk| {
            chunk.global_imports.iter().any(|import| {
                import.module == vybe_runtime::chunk::STRING_CONSTANTS_MODULE
                    && import.name == "[^\\x00-\\x7f]"
            })
        });
        let label = if selected { "candidate" } else { "frozen" };
        static SERIAL: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
        let serial = SERIAL.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
        let dir = std::path::Path::new(file!())
            .parent()
            .unwrap()
            .join("native-failures");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join(format!("html-{label}-{serial}.wasm"));
        std::fs::write(&path, &bytes).unwrap();
        let mut defined = 0usize;
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let Ok(wasmparser::Payload::CodeSectionEntry(body)) = payload {
                if body.range().contains(&error.offset()) {
                    println!(
                        "HTML {label} native failure function={defined} chunk={} file={}",
                        chunks[defined].name,
                        path.display()
                    );
                    let mut reader = body.get_operators_reader().unwrap();
                    while !reader.eof() {
                        let (op, at) = reader.read_with_offset().unwrap();
                        if at + 24 >= error.offset() && at <= error.offset() + 12 {
                            println!("HTML {label} native @{at:#x}: {op:?}");
                        }
                    }
                }
                defined += 1;
            }
        }
        panic!("HTML {label} complete native gate: {error}");
    }
    bytes.len()
}
fn program(selected: bool, base: u16) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<htmlentities-ASCII>")];
    chunks[0].local_count = base;
    globals::emit_read(&mut chunks[0], "input", 0);
    if selected {
        vybe_language_php::emitter::string_adapter::emit_htmlentities(&mut chunks, 0, 1, 0);
    } else {
        vybe_language_php::emitter::string_adapter::emit_htmlspecialchars(&mut chunks, 0, 1, 0);
        frozen::emit_named(&mut chunks[0], 0);
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn parsed(body: &str) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    let module = vybe_language_php::parse(&format!("<?php {body}")).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}
fn reference(text: &str) -> String {
    let mut out = text.to_string();
    for (from, to) in [
        ("&", "&amp;"),
        ("<", "&lt;"),
        (">", "&gt;"),
        ("\"", "&quot;"),
        ("'", "&#039;"),
    ] {
        out = out.replace(from, to);
    }
    for (from, to) in frozen::HTML_NAMED_ENTITIES_ENCODE {
        out = out.replace(from, to);
    }
    out
}
#[test]
fn exact_ascii_entities_unicode_linebreaks_scalar_and_native() {
    let old = program(false, 0);
    let new = program(true, 0);
    let old_size = validate(&old);
    let new_size = validate(&new);
    println!(
        "HTML standalone bytecode {}→{} ops {}→{} native {old_size}→{new_size}",
        old[0].code.len(),
        new[0].code.len(),
        ops(&old[0]).len(),
        ops(&new[0]).len()
    );
    let mut cases = vec![
        String::new(),
        "plain title".into(),
        "<&>\"'&copy;".into(),
        "😀�e\u{301}".into(),
    ];
    for unit in 0..=127 {
        cases.push(format!("a{}b", char::from_u32(unit).unwrap()));
    }
    for (character, _) in frozen::HTML_NAMED_ENTITIES_ENCODE {
        cases.push(format!("<{character}>&"));
    }
    for boundary in ['\u{85}', '\u{2028}', '\u{2029}'] {
        for prefix in ["", "ascii", "©"] {
            cases.push(format!("{prefix}{boundary}"));
        }
    }
    for text in cases {
        for chunks in [&old, &new] {
            let mut runtime = vm();
            runtime.set_global("input", Value::String(text.clone().into()));
            assert_eq!(
                run(&mut runtime, chunks.clone()),
                Ok(Value::String(reference(&text).into())),
                "input={text:?}"
            );
        }
    }
    for value in [
        Value::Null,
        Value::Undefined,
        Value::Bool(true),
        Value::Bool(false),
        Value::I32(123),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
    ] {
        let mut before = vm();
        let mut after = vm();
        before.set_global("input", value.clone());
        after.set_global("input", value);
        assert_eq!(run(&mut before, old.clone()), run(&mut after, new.clone()));
    }
}
#[test]
fn installed_actual_source_arguments_options_coercion_and_errors() {
    for (expression, expected) in [
        ("'plain'", "plain"),
        ("'©<&'", "&copy;&lt;&amp;"),
        ("false", ""),
        ("null", ""),
    ] {
        let old = parsed(&format!(
            "function html_subject($v) {{ return htmlentities($v); }} $out=html_subject({expression}); return $out;"
        ));
        validate(&old);
        for chunks in [old] {
            assert_eq!(run(&mut vm(), chunks), Ok(Value::String(expected.into())));
        }
    }
    let body = " $order=''; function mark($v,$label){global $order;$order.=$label;return $v;} class HtmlValue {function __toString(){global $order;$order.='s';return '©<&';}} function html_subject($v){return htmlentities(mark($v,'v'),mark(0,'1'),mark('UTF-8','2'),mark(true,'3'));} $out=html_subject(new HtmlValue()); return $order.'|'.$out;";
    let old = parsed(body);
    for chunks in [&old] {
        assert_eq!(
            run(&mut vm(), chunks.clone()),
            Ok(Value::String("v123s|&copy;&lt;&amp;".into()))
        );
    }
    validate(&old);
    let body = "class HtmlBad {function __toString(){throw new Exception('html-failure');}} function html_subject($v){return htmlentities($v);} return html_subject(new HtmlBad());";
    let old = parsed(body);
    let result = run(&mut vm(), old.clone());
    assert!(
        result.is_err(),
        "installed __toString throwing path must fail before guard"
    );
    validate(&old);
}
#[test]
fn identity_and_capacity_fallback() {
    for text in ["", "0", "a", "plain noncanonical string", "©"] {
        let input = Value::String(text.into());
        let mut observations = vec![];
        for selected in [false, true] {
            let chunks = program(selected, 0);
            let mut runtime = vm();
            runtime.set_global("input", input.clone());
            let first = run(&mut runtime, chunks.clone()).unwrap();
            let second = run(&mut runtime, chunks).unwrap();
            let (Value::String(a), Value::String(b), Value::String(original)) =
                (&first, &second, &input)
            else {
                panic!("string output")
            };
            observations.push((
                std::sync::Arc::ptr_eq(a, original),
                std::sync::Arc::ptr_eq(a, b),
            ));
        }
        assert_eq!(observations[0], observations[1], "identity {text:?}");
    }
    let budget = program(false, 0)[0].local_count;
    let old = program(false, u16::MAX - budget);
    let new = program(true, u16::MAX - budget);
    assert_eq!(old[0].local_count, u16::MAX);
    assert_eq!(new[0].local_count, u16::MAX);
    assert_eq!(old[0].code, new[0].code);
    assert_eq!(old[0].imports, new[0].imports);
    for chunks in [&old, &new] {
        let mut runtime = vm();
        runtime.set_global("input", Value::String("plain".into()));
        assert_eq!(
            run(&mut runtime, chunks.to_vec()),
            Ok(Value::String("plain".into()))
        );
    }
    for chunks in [&old, &new] {
        let error = match wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        {
            Ok(_) => panic!("expected native MAX50000 cap"),
            Err(error) => error,
        };
        assert!(error.message().contains("locals"), "{error}");
    }
    println!(
        "largest 65535-local frame: VM positive baseline/candidate proof, identical native implementation local-count cap; no native pass claimed"
    );
}

#[test]
fn installed_producer_exact_reviewed_metadata_and_native_module() {
    for base in [0u16, 1, 129] {
        for own in [false, true] {
            for prototype in [false, true] {
                let mut production = vec![Chunk::new("installed htmlentities")];
                production[0].local_count = base;
                production[0].module_instance_fields_are_own_properties = own;
                production[0].module_members_on_prototype = prototype;
                globals::emit_read(&mut production[0], "input", 37);
                let mut selected = production.clone();
                vybe_language_php::emitter::string_adapter::emit_htmlentities(
                    &mut production,
                    0,
                    1,
                    37,
                );
                candidate::emit(&mut selected, 0, 1, 37);
                for chunks in [&mut production, &mut selected] {
                    chunks[0].emit_op(Op::RETURN, 37);
                    bundle::finalize_with_runtime_helpers(chunks);
                    globals::declare_free_globals(chunks);
                    globals::normalize_global_table(chunks);
                }
                assert_eq!(production.len(), selected.len());
                for (actual, expected) in production.iter().zip(&selected) {
                    assert_eq!(actual.code, expected.code);
                    assert_eq!(actual.lines, expected.lines);
                    assert_eq!(actual.imports, expected.imports);
                    assert_eq!(actual.constants, expected.constants);
                    assert_eq!(actual.globals, expected.globals);
                    assert_eq!(actual.global_imports, expected.global_imports);
                    assert_eq!(actual.local_count, expected.local_count);
                    assert_eq!(actual.scratch_high_water, expected.scratch_high_water);
                    assert_eq!(actual.block_i32_results, expected.block_i32_results);
                    assert_eq!(actual.call_indirect_sigs, expected.call_indirect_sigs);
                    assert_eq!(actual.call_indirect_canon, expected.call_indirect_canon);
                    assert_eq!(actual.stack_switch_handlers, expected.stack_switch_handlers);
                    assert_eq!(actual.user_code_offset, expected.user_code_offset);
                    assert_eq!(
                        actual.module_instance_fields_are_own_properties,
                        expected.module_instance_fields_are_own_properties
                    );
                    assert_eq!(
                        actual.module_members_on_prototype,
                        expected.module_members_on_prototype
                    );
                }
                validate(&production);
                validate(&selected);
                assert_eq!(
                    vybe_platform_wasm::write_wasm(&production),
                    vybe_platform_wasm::write_wasm(&selected)
                );
                let options = vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: true,
                    include_name_section: true,
                };
                assert_eq!(
                    vybe_platform_wasm::write_wasm_with_options(&production, options),
                    vybe_platform_wasm::write_wasm_with_options(&selected, options)
                );
            }
        }
    }
}

#[test]
fn installed_import_capacity_keeps_legacy_emission_without_runtime_stubs() {
    // Seed every actual legacy import first. Uncalled capacity fillers are
    // declarations only; they are never registered as fake host implementations.
    let mut seed = vec![Chunk::new("legacy import seed")];
    vybe_language_php::emitter::string_adapter::emit_htmlspecialchars(&mut seed, 0, 1, 0);
    frozen::emit_named(&mut seed[0], 0);
    let legacy = seed[0].imports.clone();
    let missing = [("ecma:regexp", "test"), ("wasm:js-boolean", "cast")]
        .into_iter()
        .filter(|(m, n)| !legacy.iter().any(|i| i.module == *m && i.name == *n))
        .count();
    assert!(missing > 0);
    for (count, fallback) in [
        (usize::from(u16::MAX) + 1 - missing, false),
        (usize::from(u16::MAX) + 1, true),
    ] {
        let mut production = vec![Chunk::new("import capacity emission only")];
        for import in &legacy {
            production[0].add_import(&import.module, &import.name);
        }
        assert!(legacy
            .iter()
            .all(|import| import.module != "test:unused-capacity-declaration"));
        production[0]
            .imports
            .extend(
                (legacy.len()..count).map(|index| vybe_runtime::chunk::Import {
                    module: "test:unused-capacity-declaration".into(),
                    name: format!("unused_capacity_{index}"),
                }),
            );
        assert_eq!(production[0].imports.len(), count);
        assert_eq!(&production[0].imports[..legacy.len()], legacy.as_slice());
        globals::emit_read(&mut production[0], "input", 11);
        let before = production[0].imports.clone();
        let mut expected = production.clone();
        vybe_language_php::emitter::string_adapter::emit_htmlentities(&mut production, 0, 1, 11);
        if fallback {
            vybe_language_php::emitter::string_adapter::emit_htmlspecialchars(
                &mut expected,
                0,
                1,
                11,
            );
            frozen::emit_named(&mut expected[0], 11);
            assert_eq!(
                production[0].imports, before,
                "full import table selects legacy before mutation"
            );
        } else {
            candidate::emit(&mut expected, 0, 1, 11);
            assert_eq!(
                production[0].imports.len(),
                usize::from(u16::MAX) + 1,
                "exact largest fitting guard imports"
            );
        }
        assert_eq!(production[0].code, expected[0].code);
        assert_eq!(production[0].lines, expected[0].lines);
        assert_eq!(production[0].imports, expected[0].imports);
        assert_eq!(production[0].constants, expected[0].constants);
        assert_eq!(production[0].global_imports, expected[0].global_imports);
        assert_eq!(production[0].local_count, expected[0].local_count);
        // Do not merge unrelated bundle helper imports into an intentionally
        // full caller table. These are the complete one-caller emitted modules.
        for chunks in [&mut production, &mut expected] {
            chunks[0].emit_op(Op::RETURN, 11);
            globals::declare_free_globals(chunks);
            globals::normalize_global_table(chunks);
        }
        let bytes = vybe_platform_wasm::write_wasm(&production);
        assert_eq!(bytes, vybe_platform_wasm::write_wasm(&expected));
        let checked = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes);
        match checked {
            Ok(_)=>println!("HTML import boundary count={count} fallback={fallback}: complete native validation passed; emission-only, no host linking/VM execution claim"),
            Err(error)=>{
                let message=error.message();
                assert!(message.contains("too many types")||message.contains("too many functions")||message.contains("too many imports"),"unexpected native emission failure: {error}");
                println!("HTML import boundary count={count} fallback={fallback}: explicit native implementation cap {error}; emission-only, not native execution/pass");
            }
        }
    }
}
