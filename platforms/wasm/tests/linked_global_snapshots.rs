use std::sync::Arc;
use vybe_runtime::chunk::ConstExpr;
use vybe_runtime::{Chunk, Op, VM, Value};

fn unit(name: &str, value: i32) -> Vec<Chunk> {
    let globals = Arc::new(vec![name.to_string()].into());
    let mut entry = Chunk::new("<script>");
    entry.globals = Arc::clone(&globals);
    entry.add_global_init(name, ConstExpr::Value(Value::I32(value)));
    entry.emit_op_u32(Op::GLOBAL_GET, 0, 0);
    entry.emit_op(Op::RETURN, 0);
    let mut body = Chunk::new(&format!("body_{name}"));
    body.globals = globals;
    body.emit_op_u32(Op::GLOBAL_GET, 0, 0);
    body.emit_op(Op::RETURN, 0);
    vec![entry, body]
}

#[test]
fn later_unit_globals_survive_export_while_old_body_views_stay_immutable() {
    let mut source = VM::new();
    assert_eq!(source.run_linked(unit("first", 7), vec![]).unwrap(), Value::I32(7));
    let original_body_names = source.chunks[1].globals.clone();
    assert_eq!(source.run_linked(unit("second", 42), vec![]).unwrap(), Value::I32(42));
    assert!(source.chunks[0].globals.iter().any(|name| name == "second"));
    assert!(Arc::ptr_eq(&original_body_names, &source.chunks[1].globals));
    assert!(!original_body_names.iter().any(|name| name == "second"));

    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &source.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: true,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).expect("accumulated module must be valid Wasm");
    let decoded = vybe_platform_wasm::read_wasm(&bytes).expect("read accumulated module");
    assert!(decoded[0].globals.iter().any(|name| name == "second"));
    let mut restored = VM::new();
    assert_eq!(restored.run(decoded).unwrap(), Value::I32(7));
    assert_eq!(restored.global("first"), Some(&Value::I32(7)));
    assert_eq!(restored.global("second"), Some(&Value::I32(42)));
}
