use vybe_compiler::primitives::{bundle, globals, paths, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/path_optimization_baseline.rs"]
mod baseline;

#[derive(Clone, Copy, Debug)]
enum Operation {
    Name,
    Directory,
    Extension,
    Stem,
    HasExtension,
    ChangeExtension,
    Rooted,
    Combine,
}
const OPERATIONS: [Operation; 7] = [
    Operation::Name,
    Operation::Directory,
    Operation::Extension,
    Operation::Stem,
    Operation::HasExtension,
    Operation::ChangeExtension,
    Operation::Rooted,
];
fn program(old: bool, operation: Operation, calls: usize, argc: u8) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<path-optimization>")];
    for call in 0..calls {
        for name in ["input", "second", "third"].into_iter().take(argc as usize) {
            globals::emit_read(&mut chunks[0], name, 0);
        }
        match (old, operation) {
            (true, Operation::Name) => baseline::emit_file_name(&mut chunks[0], 0),
            (false, Operation::Name) => paths::emit_file_name(&mut chunks[0], 0),
            (true, Operation::Directory) => baseline::emit_directory(&mut chunks[0], 0),
            (false, Operation::Directory) => paths::emit_directory(&mut chunks[0], 0),
            (true, Operation::Extension) => baseline::emit_extension(&mut chunks[0], 0),
            (false, Operation::Extension) => paths::emit_extension(&mut chunks[0], 0),
            (true, Operation::Stem) => baseline::emit_file_stem(&mut chunks[0], 0),
            (false, Operation::Stem) => paths::emit_file_stem(&mut chunks[0], 0),
            (true, Operation::HasExtension) => baseline::emit_has_extension(&mut chunks[0], 0),
            (false, Operation::HasExtension) => paths::emit_has_extension(&mut chunks[0], 0),
            (true, Operation::ChangeExtension) => {
                baseline::emit_change_extension(&mut chunks[0], 0)
            }
            (false, Operation::ChangeExtension) => paths::emit_change_extension(&mut chunks[0], 0),
            (true, Operation::Rooted) => baseline::emit_is_rooted(&mut chunks[0], 0),
            (false, Operation::Rooted) => paths::emit_is_rooted(&mut chunks[0], 0),
            (true, Operation::Combine) => baseline::emit_combine(&mut chunks[0], argc, 0),
            (false, Operation::Combine) => paths::emit_combine(&mut chunks[0], argc, 0),
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
fn runtime(inputs: &[Value]) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    for (key, value) in ["input", "second", "third"].into_iter().zip(inputs) {
        vm.set_global(key, value.clone());
    }
    vm
}
fn string(s: &str) -> Value {
    Value::String(s.into())
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}

#[test]
fn numeric_path_indices_preserve_names_dots_and_utf16_results() {
    // path, name, directory, extension, stem, rooted
    let cases = [
        ("", "", "", "", "", false),
        ("file", "file", "", "", "file", false),
        ("file.txt", "file.txt", "", ".txt", "file", false),
        (".hidden", ".hidden", "", ".hidden", "", false),
        (".", ".", "", ".", "", false),
        ("..", "..", "", ".", ".", false),
        ("a.b/file", "file", "a.b", "", "file", false),
        ("a.b/file.", "file.", "a.b", ".", "file", false),
        ("/", "", "", "", "", true),
        ("a/", "", "a", "", "", false),
        (
            "C:\\dir/file.txt",
            "file.txt",
            "C:\\dir",
            ".txt",
            "file",
            true,
        ),
        (
            "\\\\host\\dir\\a.b",
            "a.b",
            "\\\\host\\dir",
            ".b",
            "a",
            true,
        ),
        ("é/😀.世", "😀.世", "é", ".世", "😀", false),
        ("😀.dir/é\\世.gz", "世.gz", "😀.dir/é", ".gz", "世", false),
    ];
    for operation in OPERATIONS {
        let argc = if matches!(operation, Operation::ChangeExtension) {
            2
        } else {
            1
        };
        let programs = [
            program(true, operation, 1, argc),
            program(false, operation, 1, argc),
        ];
        for chunks in &programs {
            validate(chunks);
        }
        for (path, name, directory, extension, stem, rooted) in cases {
            let inputs = [string(path), string("gz")];
            let before = runtime(&inputs).run(programs[0].clone()).unwrap();
            let after = runtime(&inputs).run(programs[1].clone()).unwrap();
            assert_eq!(before, after, "{operation:?} {path}");
            let expected = match operation {
                Operation::Name => string(name),
                Operation::Directory => string(directory),
                Operation::Extension => string(extension),
                Operation::Stem => string(stem),
                Operation::HasExtension => Value::Bool(!extension.is_empty()),
                Operation::Rooted => Value::Bool(rooted),
                Operation::ChangeExtension => {
                    string(&format!("{}.gz", &path[..path.len() - extension.len()]))
                }
                Operation::Combine => unreachable!(),
            };
            assert_eq!(after, expected, "{operation:?} {path}");
        }
    }
}

#[test]
fn numeric_index_specialization_keeps_existing_scalar_coercions() {
    for operation in OPERATIONS {
        let argc = if matches!(operation, Operation::ChangeExtension) {
            2
        } else {
            1
        };
        let programs = [
            program(true, operation, 1, argc),
            program(false, operation, 1, argc),
        ];
        for value in [
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::I32(12),
            Value::F64(12.5),
        ] {
            let inputs = [value, string("gz")];
            let normalize_error = |mut error: vybe_runtime::VMError| {
                // Code shrinks, so bytecode offsets change. Preserve the
                // error message/type, source lines, frame names and diagnostic.
                for frame in &mut error.call_stack {
                    frame.offset = 0;
                }
                error.to_string()
            };
            let before = runtime(&inputs)
                .run(programs[0].clone())
                .map_err(normalize_error);
            let after = runtime(&inputs)
                .run(programs[1].clone())
                .map_err(normalize_error);
            assert_eq!(
                before, after,
                "{operation:?}: preserve scalar results and strict string errors"
            );
        }
    }
}

#[test]
fn path_combination_preserves_empty_rooted_and_trailing_separator_rules() {
    for (inputs, expected) in [
        (vec![], ""),
        (vec!["a", "", "b"], "a/b"),
        (vec!["/tmp", "/etc"], "/etc"),
        (vec!["a/", "b", "c"], "a/b/c"),
        (vec!["😀/", "é", "世"], "😀/é/世"),
        (vec!["a", "C:\\foo", "bar"], "C:\\foo/bar"),
    ] {
        let inputs: Vec<_> = inputs.into_iter().map(string).collect();
        for old in [true, false] {
            let chunks = program(old, Operation::Combine, 1, inputs.len() as u8);
            validate(&chunks);
            assert_eq!(runtime(&inputs).run(chunks).unwrap(), string(expected));
        }
    }
    for extension in ["gz", ".gz", "", ".", "é", ".😀"] {
        let inputs = [string("a.b/世.txt"), string(extension)];
        let before = runtime(&inputs)
            .run(program(true, Operation::ChangeExtension, 1, 2))
            .unwrap();
        let after = runtime(&inputs)
            .run(program(false, Operation::ChangeExtension, 1, 2))
            .unwrap();
        assert_eq!(before, after);
    }
}

#[test]
fn numeric_path_recipes_reduce_repeated_emission() {
    for operation in OPERATIONS.into_iter().chain([Operation::Combine]) {
        let argc = match operation {
            Operation::ChangeExtension => 2,
            Operation::Combine => 3,
            _ => 1,
        };
        let old = program(true, operation, 20, argc);
        let new = program(false, operation, 20, argc);
        validate(&new);
        let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(bytes(&new) < bytes(&old), "{operation:?}");
        assert!(
            vybe_platform_wasm::write_wasm(&new).len() < vybe_platform_wasm::write_wasm(&old).len(),
            "{operation:?}"
        );
        println!(
            "path{operation:?}20calls: bytecode={}/{} wasm={}/{}",
            bytes(&old),
            bytes(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}

#[test]
#[ignore = "interleaved opt-level-0 path benchmark"]
fn path_execution_benchmark() {
    for operation in [Operation::Stem, Operation::HasExtension, Operation::Combine] {
        let argc = if matches!(operation, Operation::Combine) {
            3
        } else {
            1
        };
        let programs = [
            program(true, operation, 256, argc),
            program(false, operation, 256, argc),
        ];
        let inputs = [string("😀.dir/é\\世.gz"), string("é"), string("世")];
        let mut times = [vec![], vec![]];
        let mut results = [Value::Null, Value::Null];
        for round in 0..9 {
            for which in [round % 2, 1 - round % 2] {
                let mut vm = runtime(&inputs);
                let chunks = programs[which].clone();
                let start = std::time::Instant::now();
                results[which] = vm.run(chunks).unwrap();
                times[which].push(start.elapsed());
            }
        }
        assert_eq!(results[0], results[1]);
        for samples in &mut times {
            samples.sort();
        }
        println!(
            "path{operation:?}256calls: before={}us after={}us speedup={:.2}x",
            times[0][4].as_micros(),
            times[1][4].as_micros(),
            times[0][4].as_secs_f64() / times[1][4].as_secs_f64()
        );
    }
}
