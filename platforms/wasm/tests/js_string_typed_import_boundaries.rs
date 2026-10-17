use vybe_runtime::{Chunk, Op};
fn actual(name: &str, argc: u8) -> (Vec<u8>, serde_json::Value) {
    let mut c = Chunk::new(format!("<complete-{name}>"));
    c.arity = argc;
    c.param_count = argc;
    c.local_count = u16::from(argc);
    for slot in 0..argc {
        c.emit_op_u16(Op::LOCAL_GET, u16::from(slot), 19);
    }
    let i = c.add_import("wasm:js-string", name);
    c.emit_call(i, argc, 19);
    c.emit_op(Op::RETURN, 19);
    let original = c.code.clone();
    let bytes = vybe_platform_wasm::write_wasm(&[c.clone()]);
    assert_eq!(c.code, original);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let mut inventory = Vec::new();
    let mut functions = 0;
    for p in wasmparser::Parser::new(0).parse_all(&bytes) {
        match p.unwrap() {
            wasmparser::Payload::ImportSection(s) => {
                for i in s.into_imports() {
                    let i = i.unwrap();
                    let kind = match i.ty {
                        wasmparser::TypeRef::Func(_) => "function",
                        wasmparser::TypeRef::Global(_) => "global",
                        other => panic!("unexpected {other:?}"),
                    };
                    inventory
                        .push(serde_json::json!({"module":i.module,"name":i.name,"kind":kind}));
                }
            }
            wasmparser::Payload::CodeSectionEntry(_) => functions += 1,
            _ => (),
        }
    }
    assert_eq!(functions, 1);
    assert!(
        inventory
            .iter()
            .filter(|i| i["module"] == "wasm:js-string")
            .count()
            >= 18
    );
    (
        bytes,
        serde_json::json!({"name":name,"argc":argc,"completeInventory":inventory,"entireWriterModule":true}),
    )
}
#[test]
fn complete_writer_cast_and_packed_array_import_boundaries_validate() {
    for (n, a) in [
        ("cast", 1),
        ("substring", 3),
        ("fromCharCodeArray", 3),
        ("intoCharCodeArray", 3),
    ] {
        actual(n, a);
    }
}
#[test]
#[ignore = "root exports complete unmodified writer modules"]
fn export_complete_writer() {
    let p = std::path::PathBuf::from(
        std::env::var("JS_STRING_SIGNATURE_OUTPUT_DIR")
            .expect("explicit native export output directory"),
    )
    .join("actual-modules");
    std::fs::create_dir_all(&p).unwrap();
    for (n, a) in [
        ("cast", 1),
        ("substring", 3),
        ("fromCharCodeArray", 3),
        ("intoCharCodeArray", 3),
    ] {
        let (b, m) = actual(n, a);
        std::fs::write(p.join(format!("{n}.wasm")), b).unwrap();
        std::fs::write(
            p.join(format!("{n}.json")),
            serde_json::to_vec_pretty(&m).unwrap(),
        )
        .unwrap();
    }
}
