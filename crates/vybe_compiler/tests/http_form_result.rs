use std::sync::Arc;
use vybe_compiler::primitives::{bundle, functions, globals, http_form, platforms};
use vybe_runtime::{Chunk, Op, VM, Value, value::ObjectKind};

fn program() -> Vec<Chunk> {
    let mut owner = Chunk::new("form-owner");
    functions::emit_ref_func(&mut owner, 1, 0, 41);
    owner.emit_op(Op::RETURN, 41);
    let body = functions::create_function_chunk("actual-request-form", 0);
    let mut c = vec![owner, body];
    http_form::emit_parsed_body(&mut c, 1, 41);
    c[1].emit_op(Op::RETURN, 41);
    bundle::finalize_with_runtime_helpers(&mut c);
    globals::declare_free_globals(&mut c);
    globals::normalize_global_table(&mut c);
    c
}
fn field(v: &Value, key: &str) -> Value {
    let Value::Object(o) = v else {
        panic!("map object")
    };
    let o = o.lock().unwrap();
    let ObjectKind::Map(m) = &o.kind else {
        panic!("map kind")
    };
    m.iter()
        .find(|(k, _)| matches!(k,Value::String(s) if s.as_ref()==key))
        .unwrap()
        .1
        .clone()
}
fn empty(v: &Value) -> bool {
    let Value::Object(o) = v else {
        panic!("map object")
    };
    let o = o.lock().unwrap();
    let ObjectKind::Map(m) = &o.kind else {
        panic!("map kind")
    };
    m.is_empty()
}
#[test]
fn complete_request_parser_native_validation() {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&program()))
        .unwrap();
}
#[test]
fn actual_wasi_requests_both_map_branches_and_memoised_identity() {
    for (content, body, expected) in [
        (
            Some("application/x-www-form-urlencoded"),
            "name=John+Smith&op=%3D",
            Some("John Smith"),
        ),
        (
            Some("Application/X-Www-Form-Urlencoded; charset=UTF-8"),
            "name=John+Smith",
            Some("John Smith"),
        ),
        (Some("application/json"), "{\"name\":\"other\"}", None),
        (None, "", None),
    ] {
        let mut vm = VM::new();
        platforms::register_platforms_all(&mut vm);
        let f = vm.run(program()).unwrap();
        let headers = content
            .map(|s| vec![("content-type".into(), s.as_bytes().to_vec())])
            .unwrap_or_default();
        let id = vybe_platform_wasi::http::push_incoming_request(
            "POST",
            Some("/form".into()),
            Some("http".into()),
            Some("localhost".into()),
            headers,
            body.as_bytes().to_vec(),
        );
        let request = vybe_platform_wasi::http::incoming_request_value(&vm, id).unwrap();
        vm.set_global(vybe_compiler::primitives::http_request_env::REQUEST_GLOBAL, request);
        let result = vm.invoke(&f, &[]).unwrap();
        assert!(vm.last_exception.is_none());
        let fields = field(&result, "fields");
        if let Some(expected) = expected {
            assert_eq!(field(&fields, "name"), Value::String(expected.into()));
        } else {
            assert!(empty(&fields));
        }
        assert!(empty(&field(&result, "files")));
        let again = vm.invoke(&f, &[]).unwrap();
        let (Value::Object(a), Value::Object(b)) = (&result, &again) else {
            panic!("map objects")
        };
        assert!(Arc::ptr_eq(a, b), "request retains its shared parsed map");
    }
}
