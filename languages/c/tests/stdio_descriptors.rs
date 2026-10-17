use std::path::Path;
use std::sync::{Arc, Mutex};
use vybe_runtime::{Chunk, VM, Value, opcode::Op};

fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm_with_options(chunks,
        vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let mut abstract_imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                for import in section.into_imports() {
                    let import = import.unwrap();
                    if matches!(import.ty, wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)) {
                        abstract_imports.push(import.module == "ecma:value"
                            && matches!(import.name, "abstractEq" | "abstractNe"));
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                let mut reader = body.get_operators_reader().unwrap();
                while !reader.eof() {
                    if let wasmparser::Operator::Call { function_index }
                        | wasmparser::Operator::ReturnCall { function_index }
                        | wasmparser::Operator::RefFunc { function_index } = reader.read().unwrap()
                    {
                        assert_ne!(abstract_imports.get(function_index as usize), Some(&true));
                    }
                }
            }
            _ => {}
        }
    }
}

#[test]
fn fileno_standard_streams_and_file_expressions_match_native_once() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let path = root.join("tests/c/doom_startup/stdio_standard_descriptors.c");
    let directory = std::env::temp_dir().join(format!("vybe-stdio-descriptors-{}", std::process::id()));
    std::fs::create_dir_all(&directory).unwrap();
    let executable = directory.join("native");
    assert!(std::process::Command::new("cc").args(["-std=gnu11", "-O0"])
        .arg(&path).arg("-o").arg(&executable).status().unwrap().success());
    let reference = std::process::Command::new(&executable).current_dir(&directory).output().unwrap();
    assert!(reference.status.success());
    assert_eq!(reference.stdout, b"0 1 2\n");
    vybe_language_c::register();
    let mut baseline = None;
    for parser in [vybe_language_c::parse_with_peg, vybe_language_c::parse_optimized] {
        let mut bundle = vybe_compiler::projects::load(&path).unwrap();
        bundle.language.parse = parser;
        let ast = format!("{:?}", bundle.prepared_module().unwrap());
        if let Some(previous) = &baseline { assert_eq!(&ast, previous); }
        baseline = Some(ast);
        let chunks = bundle.compile().unwrap();
        validate(&chunks);
        let mut vm = VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        vm.run(chunks).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
    std::fs::remove_dir_all(directory).unwrap();
}

#[test]
fn terminal_emitter_queries_only_selected_standard_stream_and_preserves_absence() {
    for attached in [false, true] {
        for descriptor in [-1, 0, 1, 2, 3, 8] {
            let mut chunk = Chunk::new("<script>");
            chunk.emit_f64_const(f64::from(descriptor), 0);
            vybe_platform_libc::emitter::stdio_terminal::emit_isatty(&mut chunk, 0);
            chunk.emit_op(Op::RETURN, 0);
            let chunks = vec![chunk];
            validate(&chunks);
            let calls = Arc::new(Mutex::new(Vec::new()));
            let mut vm = VM::new();
            for (fd, module, getter) in [
                (0, "wasi:cli/terminal-stdin", "get-terminal-stdin"),
                (1, "wasi:cli/terminal-stdout", "get-terminal-stdout"),
                (2, "wasi:cli/terminal-stderr", "get-terminal-stderr"),
            ] {
                let calls = Arc::clone(&calls);
                vm.register_host_fn(module, getter, Box::new(move |_, _| {
                    calls.lock().unwrap().push(fd);
                    if attached {
                        Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new()))
                    } else { Value::Null }
                }));
            }
            assert_eq!(vm.run(chunks).unwrap().as_i32(), i32::from(attached && (0..=2).contains(&descriptor)));
            assert_eq!(*calls.lock().unwrap(), if (0..=2).contains(&descriptor) { vec![descriptor] } else { vec![] });
        }
    }
}
