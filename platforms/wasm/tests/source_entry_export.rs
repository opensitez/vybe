use vybe_runtime::{Chunk, Op, VM, Value};

fn exported_start_signature(bytes: &[u8]) -> (usize, usize) {
    let mut types = Vec::new();
    let mut functions = Vec::new();
    let mut start = None;
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        match payload.unwrap() {
            wasmparser::Payload::TypeSection(section) => {
                for group in section {
                    for ty in group.unwrap().into_types() {
                        types.push(match ty.composite_type.inner {
                            wasmparser::CompositeInnerType::Func(function) => {
                                Some((function.params().len(), function.results().len()))
                            }
                            _ => None,
                        });
                    }
                }
            }
            wasmparser::Payload::ImportSection(section) => {
                for import in section.into_imports() {
                    match import.unwrap().ty {
                        wasmparser::TypeRef::Func(index)
                        | wasmparser::TypeRef::FuncExact(index) => functions.push(index),
                        _ => {}
                    }
                }
            }
            wasmparser::Payload::FunctionSection(section) => {
                functions.extend(section.into_iter().map(Result::unwrap));
            }
            wasmparser::Payload::ExportSection(section) => {
                for export in section {
                    let export = export.unwrap();
                    if export.name == "_start" {
                        assert_eq!(export.kind, wasmparser::ExternalKind::Func);
                        start = Some(export.index);
                    }
                }
            }
            _ => {}
        }
    }
    types[functions[start.unwrap() as usize] as usize].unwrap()
}

fn export(chunk: Chunk) -> Vec<u8> {
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &[chunk],
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    bytes
}

#[test]
fn zero_argument_source_entry_runs_without_exposing_capture_parameter() {
    for results in [1, 2] {
        let mut root = Chunk::new("<script>");
        root.result_arity = results;
        root.globals = std::sync::Arc::new(vybe_runtime::chunk::global_index_space(
            &[],
            &[],
            &["entry_ran".into()],
        ).into());
        root.emit_i32_const(17, 0);
        root.emit_op_u32(Op::GLOBAL_SET, 3, 0);
        for _ in 0..results {
            root.emit_i32_const(0, 0);
        }
        root.emit_op(Op::RETURN, 0);
        let bytes = export(root);
        assert_eq!(exported_start_signature(&bytes), (0, 0));
        let chunks = vybe_platform_wasm::read_wasm(&bytes).unwrap();
        let ran_name = chunks[0].globals[3].clone();
        let mut vm = VM::new();
        vybe_runtime::js_builtins::register(&mut vm);
        vybe_platform_ecma::register(&mut vm);
        vm.run(chunks).unwrap();
        let globals = vm.globals_by_name();
        assert!(
            globals
                .iter()
                .any(|(name, value)| name == &ran_name && *value == Value::I32(17)),
            "{globals:?}"
        );
    }
}

#[test]
fn parameterized_source_export_is_not_implicitly_invoked() {
    let mut root = Chunk::new("<script>");
    root.arity = 1;
    root.param_count = 1;
    root.local_count = 1;
    root.emit_i32_const(0, 0);
    root.emit_op(Op::RETURN, 0);
    assert_eq!(exported_start_signature(&export(root)), (2, 1));
}

#[test]
fn declared_native_entry_keeps_its_signature() {
    let mut root = Chunk::new("<script>");
    root.func_sig = Some(("".into(), "".into()));
    root.result_arity = 0;
    root.emit_op(Op::RETURN, 0);
    assert_eq!(exported_start_signature(&export(root)), (0, 0));
}
