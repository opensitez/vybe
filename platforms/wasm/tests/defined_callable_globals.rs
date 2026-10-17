use std::sync::Arc;

use vybe_runtime::{Chunk, Op, chunk::{ConstExpr, GlobalInit, HOST_GLOBALS_MODULE, global_index_space}};

fn module(explicit_import: bool) -> Vec<Chunk> {
    let name = "module_callback";
    let hosts = if explicit_import { vec![name.to_owned()] } else { Vec::new() };
    let defined = if explicit_import { Vec::new() } else { vec![name.to_owned()] };
    let names = global_index_space(&[], &hosts, &defined);
    let index = names.iter().position(|entry| entry == name).unwrap() as u32;
    let mut root = Chunk::new("<script>");
    if explicit_import { root.add_global_import(HOST_GLOBALS_MODULE, name); }
    root.globals = Arc::new(names.into());
    root.global_inits.push(GlobalInit { name: name.into(), init: ConstExpr::RefFunc(1) });
    root.emit_op_u32(Op::GLOBAL_GET, index, 0);
    root.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    root.emit_i32_const(42, 0);
    root.emit_op(Op::I32_NE, 0);
    root.emit_if(0);
    root.emit_op(Op::UNREACHABLE, 0);
    root.emit_op(Op::END, 0);
    root.emit_i32_const(42, 0);
    root.emit_op(Op::RETURN, 0);
    let mut function = Chunk::new("callback");
    function.globals = root.globals.clone();
    function.emit_i32_const(42, 0);
    function.emit_op(Op::RETURN, 0);
    vec![root, function]
}

fn serialize(chunks: &[Chunk]) -> Vec<u8> {
    let bytes = vybe_platform_wasm::write_wasm_with_options(chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    bytes
}

fn imported_globals(bytes: &[u8]) -> Vec<(String, String, bool)> {
    let mut globals = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        if let wasmparser::Payload::ImportSection(section) = payload.unwrap() {
            for import in section.into_imports() {
                let import = import.unwrap();
                if let wasmparser::TypeRef::Global(ty) = import.ty {
                    globals.push((import.module.to_owned(), import.name.to_owned(), ty.mutable));
                }
            }
        }
    }
    globals
}

#[test]
fn module_defined_callable_initializer_does_not_create_a_host_import() {
    let bytes = serialize(&module(false));
    assert!(!imported_globals(&bytes).iter().any(|(module, name, _)|
        module == HOST_GLOBALS_MODULE && name == "module_callback"));
    let decoded = vybe_platform_wasm::read_wasm(&bytes).unwrap();
    assert!(!decoded.iter().flat_map(|chunk| &chunk.global_imports)
        .any(|import| import.module == HOST_GLOBALS_MODULE && import.name == "module_callback"));
}

#[test]
fn module_defined_callable_executes_after_metadata_free_reload() {
    let bytes = serialize(&module(false));
    let decoded = vybe_platform_wasm::read_wasm(&bytes).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    vybe_platform_ecma::register(&mut vm);
    // The reader's _start adapter discards results. The module itself traps
    // unless its initialized callable executes and returns the expected value.
    vm.run(decoded).unwrap();
}

#[test]
fn reloaded_start_really_checks_the_initialized_callable_result() {
    let mut chunks = module(false);
    chunks[1].code.clear();
    chunks[1].emit_i32_const(41, 0);
    chunks[1].emit_op(Op::RETURN, 0);
    let decoded = vybe_platform_wasm::read_wasm(&serialize(&chunks)).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    vybe_platform_ecma::register(&mut vm);
    let error = vm.run(decoded).unwrap_err();
    assert!(error.to_string().contains("unreachable"), "{error}");
}

#[test]
fn explicit_callable_global_import_keeps_its_binding_and_mutability() {
    let bytes = serialize(&module(true));
    assert!(imported_globals(&bytes).contains(&(
        HOST_GLOBALS_MODULE.into(), "module_callback".into(), true)));
}
