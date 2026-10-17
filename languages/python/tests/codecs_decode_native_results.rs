use vybe_compiler::primitives;
use vybe_runtime::{Chunk,Op,VM,Value};
fn snapshot(value: &Value) -> serde_json::Value {
    match value {
        Value::Bool(b) => serde_json::json!(b),
        Value::I32(n) => serde_json::json!(n),
        Value::F64(n) => serde_json::json!(n),
        Value::String(s) => serde_json::json!(s.as_ref()),
        Value::Null => serde_json::Value::Null,
        Value::Object(o) => {
            let object = o.lock().unwrap();
            match &object.kind {
                vybe_runtime::value::ObjectKind::Array(a) => {
                    serde_json::json!({"kind":"array","items":a.iter().map(snapshot).collect::<Vec<_>>()})
                }
                vybe_runtime::value::ObjectKind::TypedArray(_) => {
                    drop(object);
                    serde_json::json!({"kind":"typed","bytes":vybe_platform_ecma::typedarray::bytes_from_value(value)})
                }
                _ => panic!("unexpected result kind"),
            }
        }
        _ => panic!("unexpected {value:?}"),
    }
}

#[test]
fn codecs_decode_branches_produce_native_results(){
 vybe_language_python::register();
 for(codec,input,expected) in [("hex","436174",serde_json::json!({"kind":"typed","bytes":[67,97,116]})),("base64","Q2F0",serde_json::json!({"kind":"typed","bytes":[67,97,116]})),("rot_13","Png",serde_json::json!("Cat")),("ascii","Cat",serde_json::json!("Cat")),("utf-8","Cat",serde_json::json!("Cat")),("utf-16","Cat",serde_json::json!("Cat"))]{
  let mut c=Chunk::new("codec");c.emit_string_const(input,0);c.emit_string_const(codec,0);let mut chunks=vec![c];vybe_language_python::emitter::base64_adapter::emit_codecs_decode(&mut chunks,0,2,0);chunks[0].emit_op(Op::RETURN,0);
  primitives::bundle::finalize_with_runtime_helpers(&mut chunks);primitives::globals::declare_free_globals(&mut chunks);primitives::globals::normalize_global_table(&mut chunks);
  #[cfg(feature="native-gates")]
  wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&chunks)).unwrap();
  let mut vm=VM::new();primitives::platforms::register_platforms_all(&mut vm);assert_eq!(snapshot(&vm.run(chunks).unwrap()),expected,"{codec}");
 }
}
