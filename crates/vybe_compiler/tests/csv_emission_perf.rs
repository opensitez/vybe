//! CSV producer specialization and shared-body regressions.
use vybe_compiler::primitives::{bundle, csv, globals, platforms};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
use vybex as _;
#[path = "fixtures/dynamic_csv.rs"]
mod baseline;

fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn vm(input: Value, delimiter: &str, enclosure: &str) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm.set_global("delimiter", Value::String(delimiter.into()));
    vm.set_global("enclosure", Value::String(enclosure.into()));
    vm
}
fn finish(mut chunks: Vec<Chunk>) -> Vec<Chunk> {
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn program(old: bool, operation: usize, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<csv>")];
    for i in 0..calls {
        for name in ["input", "delimiter", "enclosure"] {
            globals::emit_read(&mut chunks[0], name, 0);
        }
        match (old, operation) {
            (true, 0) => baseline::emit_parse_line(&mut chunks, 0, 0),
            (false, 0) => csv::emit_parse_line(&mut chunks, 0, 0),
            (true, 1) => baseline::emit_parse_document(&mut chunks, 0, 0),
            (false, 1) => csv::emit_parse_document(&mut chunks, 0, 0),
            (true, 2) => {
                baseline::emit_format_row(&mut chunks, 0, baseline::FormatOptions::minimal(), 0)
            }
            (false, 2) => csv::emit_format_row(&mut chunks, 0, csv::FormatOptions::minimal(), 0),
            (true, 3) => baseline::emit_format_row(
                &mut chunks,
                0,
                baseline::FormatOptions::quote_whitespace(),
                0,
            ),
            (false, 3) => {
                csv::emit_format_row(&mut chunks, 0, csv::FormatOptions::quote_whitespace(), 0)
            }
            _ => unreachable!(),
        }
        if i + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    finish(chunks)
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn normalized(value: Value) -> String {
    if let Value::Object(object) = &value {
        if let ObjectKind::Array(values) = &object.lock().unwrap().kind {
            return format!(
                "[{}]",
                values
                    .iter()
                    .cloned()
                    .map(normalized)
                    .collect::<Vec<_>>()
                    .join(",")
            );
        }
    }
    format!("{value:?}")
}
fn fields(values: &[&str]) -> Value {
    array(values.iter().map(|s| Value::String((*s).into())).collect())
}
#[test]
fn line_parser_keeps_runtime_dialects_quotes_unicode_and_empty_fields() {
    for (input, delimiter, enclosure, expected) in [
        ("", ",", "\"", vec![""]),
        ("a,,b,", ",", "\"", vec!["a", "", "b", ""]),
        (
            "\"a,b\",\"q\"\"q\",last",
            ",",
            "\"",
            vec!["a,b", "q\"q", "last"],
        ),
        ("'a;b';'';'é😀'", ";", "'", vec!["a;b", "", "é😀"]),
        ("\"a\nb\",c", ",", "\"", vec!["a\nb", "c"]),
        ("😀,x", ",", "", vec!["😀", "x"]),
        ("😀,x", "", "", vec!["😀,x"]),
        ("😀�x�,y", ",", "�", vec!["😀x", "y"]),
        ("😀�x", "�", "\"", vec!["😀", "x"]),
    ] {
        for old in [true, false] {
            if old && input.contains("😀") {
                continue;
            }
            let chunks = program(old, 0, 1);
            if !old {
                validate(&chunks);
            }
            assert_eq!(
                normalized(
                    vm(Value::String(input.into()), delimiter, enclosure)
                        .run(chunks)
                        .unwrap()
                ),
                normalized(fields(&expected))
            );
        }
    }
}
#[test]
fn document_parser_keeps_record_boundaries_and_quoted_empty_records() {
    for (input, expected) in [
        ("", vec![]),
        ("\"\"", vec![vec![""]]),
        ("\"\"\r\n", vec![vec![""]]),
        ("a,b\r\nc,d\n", vec![vec!["a", "b"], vec!["c", "d"]]),
        (
            "\"a\nb\",c\rd,\"q\"\"q\"",
            vec![vec!["a\nb", "c"], vec!["d", "q\"q"]],
        ),
        ("\n\r\n", vec![vec![""], vec![""]]),
    ] {
        for old in [true, false] {
            let chunks = program(old, 1, 1);
            if !old {
                validate(&chunks);
            }
            assert_eq!(
                normalized(
                    vm(Value::String(input.into()), ",", "\"")
                        .run(chunks)
                        .unwrap()
                ),
                normalized(array(expected.iter().map(|row| fields(row)).collect())),
                "{input:?}"
            );
        }
    }
}
#[test]
fn formatter_preserves_declared_whitespace_policy_and_dialect() {
    for operation in [2, 3] {
        for (delimiter, enclosure) in [(",", "\""), (";", "'")] {
            let input = fields(&[
                "",
                "has space",
                "a,b",
                "q\"q",
                "a;b",
                "q'q",
                "é😀",
                "line\nbreak",
            ]);
            let old = program(true, operation, 1);
            let new = program(false, operation, 1);
            validate(&new);
            let expected = vm(input.clone(), delimiter, enclosure).run(old).unwrap();
            let actual = vm(input, delimiter, enclosure).run(new).unwrap();
            assert_eq!(actual.as_str(), expected.as_str());
            if operation == 3 {
                assert!(
                    actual
                        .as_str()
                        .contains(&format!("{enclosure}has space{enclosure}"))
                );
            } else {
                assert!(actual.as_str().contains("has space"));
            }
        }
    }
}
#[test]
fn shared_csv_calls_finalize_once_and_shrink_repeated_emission() {
    for operation in 0..4 {
        let old = program(true, operation, 20);
        let mut new = program(false, operation, 20);
        let codes: Vec<_> = new.iter().map(|c| c.code.clone()).collect();
        bundle::finalize_with_runtime_helpers(&mut new);
        assert_eq!(
            new.iter().map(|c| c.code.clone()).collect::<Vec<_>>(),
            codes
        );
        validate(&new);
        let before: usize = old.iter().map(|c| c.code.len()).sum();
        let after: usize = new.iter().map(|c| c.code.len()).sum();
        println!(
            "csv operation={operation}: bytecode={before}/{after} wasm={}/{}",
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
        assert!(after * 8 < before);
        let input = if operation < 2 {
            Value::String("a,\"b,c\",d".into())
        } else {
            fields(&["a", "b,c", "d"])
        };
        assert_eq!(
            normalized(vm(input.clone(), ",", "\"").run(new).unwrap()),
            normalized(vm(input, ",", "\"").run(old).unwrap())
        );
    }
}
#[test]
fn php_csv_frontend_keeps_dynamic_dialect_and_scalar_coercion() {
    vybe_language_php::register();
    for (body, expected) in [
        (
            "return str_getcsv('a;b;\"c;d\"', ';');",
            vec!["a", "b", "c;d"],
        ),
        ("return str_getcsv('a,b');", vec!["a", "b"]),
        ("return str_getcsv(42);", vec!["42"]),
        ("$d = ';'; return str_getcsv('a;b', $d);", vec!["a", "b"]),
        (
            "function sep() { return ';'; } return str_getcsv('a;b', sep());",
            vec!["a", "b"],
        ),
        ("$s = '\"😀\",x'; return str_getcsv($s);", vec!["😀", "x"]),
    ] {
        let module = vybe_language_php::parse(&format!("<?php {body}")).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        validate(&chunks);
        let actual = normalized(vm(Value::Null, ",", "\"").run(chunks.clone()).unwrap());
        let expected = normalized(fields(&expected));
        if actual != expected {
            std::fs::write(
                std::env::temp_dir().join("csv-frontend-failure.txt"),
                chunks
                    .iter()
                    .map(vybe_runtime::debug::disassemble)
                    .collect::<Vec<_>>()
                    .join("\n"),
            )
            .unwrap();
        }
        assert_eq!(actual, expected, "{body}");
    }
}
#[test]
#[ignore = "manual interleaved debug execution benchmark"]
fn csv_execution_benchmark() {
    for operation in 0..4 {
        let input = if operation == 0 {
            Value::String(
                (0..128)
                    .map(|i| format!("\"item{i},é\""))
                    .collect::<Vec<_>>()
                    .join(",")
                    .into(),
            )
        } else if operation == 1 {
            Value::String("\"a\nb\",c,d\r\n".repeat(128).into())
        } else {
            array(
                (0..128)
                    .map(|i| Value::String(format!("field {i},é").into()))
                    .collect(),
            )
        };
        let old = program(true, operation, 1);
        let new = program(false, operation, 1);
        let expected = normalized(vm(input.clone(), ",", "\"").run(old.clone()).unwrap());
        let mut times = [Vec::new(), Vec::new()];
        for repeat in 0..5 {
            for which in if repeat % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut vm = vm(input.clone(), ",", "\"");
                let chunks = if which == 0 { old.clone() } else { new.clone() };
                let start = std::time::Instant::now();
                let result = vm.run(chunks).unwrap();
                times[which].push(start.elapsed().as_micros());
                assert_eq!(normalized(result), expected);
            }
        }
        for t in &mut times {
            t.sort();
        }
        println!(
            "csv operation={operation}: dynamic={}us specialized={}us speedup={:.2}x",
            times[0][2],
            times[1][2],
            times[0][2] as f64 / times[1][2] as f64
        );
    }
}

#[test]
#[ignore = "manual interleaved long-field execution benchmark"]
fn csv_plain_field_run_benchmark() {
    for size in [256, 8192, 16384] {
        let input = Value::String("a".repeat(size).into());
        let old = program(true, 0, 1);
        let new = program(false, 0, 1);
        let expected = normalized(vm(input.clone(), ",", "\"").run(old.clone()).unwrap());
        let mut timings = [Vec::new(), Vec::new()];
        for repeat in 0..5 {
            for which in if repeat % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut vm = vm(input.clone(), ",", "\"");
                let chunks = if which == 0 { old.clone() } else { new.clone() };
                let start = std::time::Instant::now();
                let result = vm.run(chunks).unwrap();
                timings[which].push(start.elapsed().as_micros());
                assert_eq!(normalized(result), expected);
            }
        }
        for times in &mut timings {
            times.sort();
        }
        println!(
            "csv plain size={size}: dynamic={}us runs={}us speedup={:.2}x",
            timings[0][2],
            timings[1][2],
            timings[0][2] as f64 / timings[1][2] as f64
        );
    }
}

#[test]
fn csv_round_trip_preserves_unicode_and_literal_dollar_enclosures() {
    for (delimiter, enclosure) in [(",", "\""), (";", "'"), ("|", "$")] {
        let input = fields(&[
            "",
            "😀",
            "é😀tail",
            "a,b;c|d",
            "q\"q'q$q",
            "line\r\nbreak",
            "last",
        ]);
        for operation in [2, 3] {
            let format = program(false, operation, 1);
            let parse = program(false, 0, 1);
            validate(&format);
            validate(&parse);
            let rendered = vm(input.clone(), delimiter, enclosure).run(format).unwrap();
            if enclosure == "$" {
                assert!(rendered.as_str().contains("$$"));
            }
            let output = vm(rendered, delimiter, enclosure).run(parse).unwrap();
            assert_eq!(normalized(output), normalized(input.clone()));
        }
    }
}

#[test]
fn token_cache_keeps_dense_and_sparse_document_boundaries() {
    for delimiter in [",", ";", "|"] {
        let row = ["a", "", "😀", "q\"q", "long ordinary tail", "last"];
        let rendered = vm(fields(&row), delimiter, "\"")
            .run(program(false, 2, 1))
            .unwrap()
            .as_str()
            .to_owned();
        let document = vec![rendered; 64].join("\r\n");
        let chunks = program(false, 1, 1);
        validate(&chunks);
        let actual = vm(Value::String(document.into()), delimiter, "\"")
            .run(chunks)
            .unwrap();
        assert_eq!(
            normalized(actual),
            normalized(array((0..64).map(|_| fields(&row)).collect()))
        );
    }
}

#[test]
#[ignore = "manual export for independent wasm-tools validation"]
fn export_csv_wasm_fixtures() {
    let dir = std::env::temp_dir().join("vybe-csv-wasm");
    std::fs::create_dir_all(&dir).unwrap();
    for operation in 0..4 {
        for calls in [1, 20] {
            std::fs::write(
                dir.join(format!("operation-{operation}-calls-{calls}.wasm")),
                vybe_platform_wasm::write_wasm(&program(false, operation, calls)),
            )
            .unwrap();
        }
    }
}

#[test]
fn php_empty_csv_is_null_field_in_folded_and_runtime_paths() {
    vybe_language_php::register();
    for body in [
        "return str_getcsv('');",
        "$s = ''; return str_getcsv($s);",
        "$s = null; return str_getcsv($s);",
        "function empty_text() { return ''; } return str_getcsv(empty_text());",
    ] {
        let module = vybe_language_php::parse(&format!("<?php {body}")).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        validate(&chunks);
        let actual = vm(Value::Null, ",", "\"").run(chunks).unwrap();
        assert_eq!(
            normalized(actual),
            normalized(array(vec![Value::Null])),
            "{body}"
        );
    }
}
