use vybe_compiler::primitives::{platforms, strings};
use vybe_runtime::{Chunk, VM, Value};
fn source_program(source: &str, selected: bool) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_python::register);
    let mut module = vybe_language_python::parse(source).unwrap();
    module
        .body
        .push(vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
            vybe_ast::Expression::ident("answer"),
        ))));
    let profile = vybe_language_python::profile_source();
    let original = profile.replace(
        "format_map      = { emit = \"intrinsic:format_map\",",
        "format_map      = { emit = \"common:python.format_map\",",
    );
    let profile =
        vybe_compiler::profile::parse_profile(if selected { profile } else { &original }).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}
fn vm() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}
fn validate(chunks: &[Chunk], label: &str) {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        eprintln!("literal validation label={label}");
        let mut names = std::collections::BTreeMap::new();
        let named_bytes = vybe_platform_wasm::write_wasm_with_options(
            chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_name_section: true,
                ..Default::default()
            },
        );
        for payload in wasmparser::Parser::new(0).parse_all(&named_bytes) {
            if let wasmparser::Payload::CustomSection(section) = payload.unwrap() {
                if let wasmparser::KnownCustom::Name(reader) = section.as_known() {
                    for name in reader {
                        if let wasmparser::Name::Function(map) = name.unwrap() {
                            for entry in map {
                                let entry = entry.unwrap();
                                names.insert(entry.index, entry.name.to_string());
                            }
                        }
                    }
                }
            }
        }
        let mut first_defined = 0;
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let wasmparser::Payload::ImportSection(section) = payload.unwrap() {
                for import in section.into_imports() {
                    if matches!(
                        import.unwrap().ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                    ) {
                        first_defined += 1;
                    }
                }
            }
        }
        let mut function = first_defined;
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
                if body.range().contains(&error.offset()) {
                    eprintln!(
                        "native failure function={} name={:?} range={:?} error={error}",
                        function,
                        names.get(&function),
                        body.range()
                    );
                    let mut operators = body.get_operators_reader().unwrap();
                    while !operators.eof() {
                        let (operator, offset) = operators.read_with_offset().unwrap();
                        if offset + 100 >= error.offset() && offset <= error.offset() + 30 {
                            eprintln!("  wasm@{offset}: {operator:?}");
                            if let wasmparser::Operator::Call { function_index } = operator {
                                eprintln!("    target={:?}", names.get(&function_index));
                            }
                        }
                    }
                }
                function += 1;
            }
        }
        panic!("native module rejected [{label}]: {error}");
    }
}

fn logs(runtime: &VM) -> Vec<String> {
    let Some(Value::Object(object)) = runtime.global("events") else {
        return vec![];
    };
    let object = object.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(values) = &object.kind else {
        panic!("events must be array")
    };
    values.iter().map(|v| v.as_str().to_owned()).collect()
}
fn result(runtime: &mut VM, chunks: Vec<Chunk>) -> Result<Value, String> {
    runtime.run(chunks).map_err(|e| e.message)
}
#[test]
fn literal_gate_matches_existing_utf16_scanner_without_escape_reinterpretation() {
    for (text, eligible) in [
        ("{a}".repeat(63), false),
        ("{a}".repeat(64), true),
        ("{{a}}".repeat(64), true),
        ("😀{a}".repeat(64), true),
        ("{a".repeat(64), false),
        ("}".repeat(192), false),
    ] {
        assert_eq!(
            strings::format_map_literal_uses_large_path(&text),
            eligible,
            "{text:?}"
        );
    }
}
#[test]
fn parsed_source_selects_only_many_placeholder_literals_and_preserves_outputs() {
    let template = "{a}".repeat(64);
    for (source, inline, expected) in [
        (
            "answer = '{a}:{a}'.format_map({'a':'x'})\n".to_owned(),
            false,
            "x:x".to_owned(),
        ),
        (
            format!("template={template:?}\nanswer=template.format_map({{'a':'x'}})\n"),
            false,
            "x".repeat(64),
        ),
        (
            format!("answer={template:?}.format_map({{'a':'x'*65536}})\n"),
            true,
            "x".repeat(64 * 65536),
        ),
        (
            format!("answer={template:?}.format_map({{'a':'small'}})\n"),
            true,
            "small".repeat(64),
        ),
        (
            format!(
                "answer={:?}.format_map({{'a':'é😀'}})\n",
                "é😀{a}".repeat(64)
            ),
            true,
            "é��é😀".repeat(64),
        ),
    ] {
        let old = source_program(&source, false);
        let new = source_program(&source, true);
        validate(&old, &format!("baseline-common source={source}"));
        validate(&new, &format!("selected-intrinsic source={source}"));
        assert_eq!(
            new.iter().any(|c| c.name == "__stdlib_format_map"),
            !inline,
            "source={source}"
        );
        assert_eq!(
            old.iter()
                .filter(|c| c.name == "__stdlib_format_map")
                .count(),
            1
        );
        let before = result(&mut vm(), old);
        let after = result(&mut vm(), new);
        assert_eq!(before, after, "source={source}");
        assert_eq!(after.unwrap(), Value::String(expected.into()));
    }
}
#[test]
fn source_evaluation_aliases_shadowing_and_errors_preserve_typed_contract() {
    let template = "{a}".repeat(64);
    let cases=[
        (format!("events=[]\ndef mapping():\n    events.append('mapping')\n    return {{'a':'x'*65536}}\nanswer={template:?}.format_map(mapping())\n"),vec!["mapping"],Some("x".repeat(64*65536))),
        (format!("events=[]\ndef receiver():\n    events.append('receiver')\n    return {template:?}\ndef mapping():\n    events.append('mapping')\n    return {{'a':'x'}}\nanswer=receiver().format_map(mapping())\n"),vec!["receiver","mapping"],Some("x".repeat(64))),
        ("class Custom:\n    def format_map(self, mapping):\n        return 'custom'\nanswer=Custom().format_map({})\n".to_owned(),vec![],Some("custom".to_owned())),
        // Existing first-class helper binding yields undefined in both profiles.
        ("alias='{a}'.format_map\nanswer=alias({'a':'alias'})\n".to_owned(),vec![],Some("undefined".to_owned())),
        (format!("answer={template:?}.format_map(None)\n"),vec![],Some("null".repeat(64))),
        (format!("def mapping():\n    raise ValueError('mapping failed')\nanswer={template:?}.format_map(mapping())\n"),vec![],None),
    ];
    for (source, expected_logs, expected) in cases {
        let mut before = vm();
        let mut after = vm();
        let old = source_program(&source, false);
        let new = source_program(&source, true);
        validate(&old, &format!("baseline-common source={source}"));
        validate(&new, &format!("selected-intrinsic source={source}"));
        let old = result(&mut before, old);
        let new = result(&mut after, new);
        assert_eq!(old, new, "source={source}");
        assert_eq!(logs(&before), expected_logs);
        assert_eq!(logs(&after), expected_logs);
        if let Some(expected) = expected {
            assert_eq!(new.unwrap(), Value::String(expected.into()));
        } else {
            assert!(new.is_err());
        }
    }
}
#[test]
fn selected_source_own_lookup_conversion_callbacks_and_partial_errors_match() {
    let template = "{a}".repeat(64);
    for fail in [false, true] {
        let failure = if fail {
            "    if calls == 3:\n        raise ValueError('conversion failed')\n"
        } else {
            ""
        };
        // ARRAY_GET supplies the receiver slot plus key; a raw free callback
        // declares only the explicit key under the module receiver ABI.
        let source = format!(
            "events=[]\ncalls=0\nclass Box:\n    pass\ndef convert():\n    global calls\n    calls += 1\n    events.append('convert')\n{failure}    return 'x'*65536\nvalue=Box()\nsetattr(value, 'toString', convert)\ndef lookup(key):\n    events.append('get:'+key)\n    return value\nmapping=Box()\nsetattr(mapping, 'a', None)\nsetattr(mapping, '__getitem__', lookup)\nanswer={template:?}.format_map(mapping)\n"
        );
        let mut old_vm = vm();
        let mut new_vm = vm();
        let old = source_program(&source, false);
        let new = source_program(&source, true);
        validate(&old, &format!("baseline-common source={source}"));
        validate(&new, &format!("selected-intrinsic source={source}"));
        assert!(!new.iter().any(|c| c.name == "__stdlib_format_map"));
        let before = result(&mut old_vm, old);
        let after = result(&mut new_vm, new);
        for (label, runtime, output) in [("baseline", &old_vm, &before), ("selected", &new_vm, &after)] {
            let result_summary = match output {
                Ok(Value::String(text)) => format!("String {} bytes", text.len()),
                Ok(value) => format!("Value tag={:?}", value.tag()),
                Err(message) => format!("Error {message}"),
            };
            println!("callback source {label} fail={fail} output={result_summary} trace={:?}", logs(runtime));
            for name in ["mapping", "value"] {
                if let Some(Value::Object(object)) = runtime.global(name) {
                    let object = object.lock().unwrap();
                    let mut keys = object.properties.keys().collect::<Vec<_>>();
                    keys.sort();
                    let field_tags = object.fields.iter().map(Value::tag).collect::<Vec<_>>();
                    println!("  {name} type={} own_keys={keys:?} field_tags={field_tags:?}", object.type_id);
                    for key in ["__getitem__", "toString", "a"] {
                        if let Some(value) = object.properties.get(key) {
                            println!("    property {key} tag={:?}", value.tag());
                            if let Value::Object(callback) = value {
                                let callback = callback.lock().unwrap();
                                if let vybe_runtime::value::ObjectKind::Function(function) = &callback.kind {
                                    println!("    function name={:?} arity={} chunk_index={}", function.name, function.arity, function.chunk_index);
                                }
                                for slot in ["__table_idx", "__vybe_method_receiver", "__vybe_method_target"] {
                                    if let Some(value) = callback.properties.get(slot) {
                                        println!("    callable slot {slot} tag={:?} numeric={:?}", value.tag(), matches!(value, Value::I32(_) | Value::F64(_)).then(|| value.as_f64()));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        assert_eq!(before, after, "fail={fail}");
        let expected: Vec<String> = (0..if fail { 3 } else { 64 })
            .flat_map(|_| ["get:a".to_owned(), "convert".to_owned()])
            .collect();
        assert_eq!(logs(&old_vm), expected);
        assert_eq!(logs(&new_vm), expected);
        if fail {
            assert!(after.unwrap_err().contains("conversion failed"));
        } else {
            assert_eq!(after.unwrap(), Value::String("x".repeat(64 * 65536).into()));
        }
    }
}
#[test]
fn selected_source_declared_and_inherited_protocol_policy_matches_baseline() {
    let template = "{a}".repeat(64);
    for inherited in [false, true] {
        let declaration = if inherited {
            "class MappingBase:\n    def __getitem__(self, key):\n        events.append('declared-get')\n        return value\nclass Mapping(MappingBase):\n    pass\nclass ValueBase:\n    def __str__(self):\n        events.append('declared-str')\n        return 'x'*65536\nclass Render(ValueBase):\n    pass\n"
        } else {
            "class Mapping:\n    def __getitem__(self, key):\n        events.append('declared-get')\n        return value\nclass Render:\n    def __str__(self):\n        events.append('declared-str')\n        return 'x'*65536\n"
        };
        let source = format!(
            "events=[]\n{declaration}value=Render()\nmapping=Mapping()\nmapping.a=None\nanswer={template:?}.format_map(mapping)\n"
        );
        let mut before = vm();
        let mut after = vm();
        let old = source_program(&source, false);
        let new = source_program(&source, true);
        validate(&old, &format!("baseline-common source={source}"));
        validate(&new, &format!("selected-intrinsic source={source}"));
        let old_result = result(&mut before, old);
        let new_result = result(&mut after, new);
        assert_eq!(old_result, new_result, "inherited={inherited}");
        let baseline_trace = logs(&before);
        assert_eq!(baseline_trace, logs(&after), "inherited={inherited}");
        // Capture the actual role-lookup contract, including existing gaps;
        // the positive own-function case above separately requires64 callbacks.
        println!("format_map declared protocol inherited={inherited} trace={baseline_trace:?}");
    }
}
#[test]
#[ignore = "parsed-source opt-level-0 cold cost and repeated inline emission experiment"]
fn benchmark_literal_selection_cold_and_repeated_calls() {
    use std::time::Instant;
    for (name, template, size, calls) in [
        ("short192k", "{a}{a}{a}".to_owned(), 65536, 1),
        ("unknown192k", "{a}{a}{a}".to_owned(), 65536, 16),
        ("large64x64k", "{a}".repeat(64), 65536, 1),
        ("large256x64k", "{a}".repeat(256), 65536, 1),
        ("repeat-literal16", "{a}".repeat(64), 2, 16),
    ] {
        let expression = if name.starts_with("unknown") {
            "template.format_map(data)".to_owned()
        } else {
            format!("{template:?}.format_map(data)")
        };
        let source = format!(
            "template={template:?}\ndata={{'a':'x'*{size}}}\n{}",
            format!("answer={expression}\n").repeat(calls)
        );
        let programs = [
            source_program(&source, false),
            source_program(&source, true),
        ];
        for (index, program) in programs.iter().enumerate() {
            validate(program, &format!("benchmark={name} variant={index}"));
        }
        let mut times: [Vec<u128>; 2] = std::array::from_fn(|_| vec![]);
        let expected = Value::String("x".repeat(template.matches("{a}").count() * size).into());
        for sample in 0..7 {
            for previous in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut runtime = vm();
                let prepared = programs[previous].clone();
                let start = Instant::now();
                let actual = runtime.run(prepared).unwrap();
                let elapsed = start.elapsed();
                assert_eq!(actual, expected);
                times[previous].push(elapsed.as_nanos());
            }
        }
        for list in &mut times {
            list.sort();
        }
        let bytecode = |cs: &[Chunk]| cs.iter().map(|c| c.code.len()).sum::<usize>();
        println!(
            "literal format_map {name} calls={calls} bytecode={}=>{} wasm={}=>{} old={}ns new={}ns gain={:.2}x",
            bytecode(&programs[0]),
            bytecode(&programs[1]),
            vybe_platform_wasm::write_wasm(&programs[0]).len(),
            vybe_platform_wasm::write_wasm(&programs[1]).len(),
            times[0][3],
            times[1][3],
            times[0][3] as f64 / times[1][3] as f64
        );
    }
}
