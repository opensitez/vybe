use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM};
#[test]
fn measure_catalog_helpers() {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let names: Vec<String> =
        serde_json::from_slice(&std::fs::read(dir.join("HELPER_SIZE_NAMES.json")).unwrap())
            .unwrap();
    let original_hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(|_| {}));
    let mut rows = vec![];
    for (index, name) in names.iter().enumerate() {
        let mut row = serde_json::json!({"name":name});
        let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let mut imports = Chunk::new("imports");
            let name: &'static str = Box::leak(name.clone().into_boxed_str());
            let c = primitives::polyfills::build_runtime_helpers_for_exports(&mut imports, &[name])
                .chunks
                .remove(0);
            row["body_bytes"] = c.code.len().into();
            row["locals"] = c.local_count.into();
            row["imports"] = c.imports.len().into();
            row["arity"] = c.arity.into();
            let mut root = Chunk::new("helper-size-owner");
            primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
            root.emit_op(Op::RETURN, 0);
            let mut cs = vec![root, c];
            primitives::bundle::finalize_with_runtime_helpers(&mut cs);
            primitives::globals::declare_free_globals(&mut cs);
            primitives::globals::normalize_global_table(&mut cs);
            row["linked_body_bytes"] = cs.iter().map(|c| c.code.len()).sum::<usize>().into();
            row["linked_chunks"] = cs.len().into();
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            row["native_bytes"] = bytes.len().into();
            row["validation"] =
                match wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&bytes)
                {
                    Ok(_) => serde_json::json!("valid"),
                    Err(e) => serde_json::json!(e.to_string()),
                };
        }));
        if let Err(e) = outcome {
            let reason = if let Some(s) = e.downcast_ref::<String>() {
                s.clone()
            } else if let Some(s) = e.downcast_ref::<&str>() {
                (*s).into()
            } else {
                "non-text panic".into()
            };
            row["build_error"] = reason.into();
        }
        rows.push(row);
        if index % 50 == 0 {
            eprintln!("Measured {} / {} helper names", index + 1, names.len());
        }
    }
    std::panic::set_hook(original_hook);
    rows.sort_by_key(|r| std::cmp::Reverse(r["body_bytes"].as_u64().unwrap_or(0)));
    let result = serde_json::json!({"scope":"Named helper catalog from build_runtime_helper_export. Size ranking, not execution profile; includes legacy catalog entries without live callers. Native validation failures and unavailable builders recorded explicitly.","names":names.len(),"built":rows.iter().filter(|r|r.get("body_bytes").is_some()).count(),"valid_modules":rows.iter().filter(|r|r["validation"]=="valid").count(),"rows":rows});
    std::fs::write(
        dir.join("HELPER_SIZE_AUDIT.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!(
        "catalog names={} built={} valid={}",
        result["names"], result["built"], result["valid_modules"]
    );
    assert!(result["built"].as_u64().unwrap() > 50);
}
