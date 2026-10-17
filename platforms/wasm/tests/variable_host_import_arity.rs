use vybe_runtime::{Chunk, Op, opcode::heaptype::HT_EXTERN};

#[test]
fn same_host_name_with_two_call_arities_keeps_both_signatures() {
    let mut chunk = Chunk::new("<script>");
    let push = chunk.add_import("ecma:array", "push");
    for argc in [2, 3, 65, 255] {
        for _ in 0..argc {
            chunk.emit_ref_null(HT_EXTERN, 0);
        }
        chunk.emit_call(push, argc, 0);
        chunk.emit_op(Op::DROP, 0);
    }
    chunk.emit_ref_null(HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();

    let mut signatures = Vec::new();
    let mut padding_globals = 0;
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(imports) => {
                for import in imports.into_imports() {
                    let import = import.unwrap();
                    if import.module == "ecma:array" && import.name == "push" {
                        signatures.push(import.ty);
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                for op in body.get_operators_reader().unwrap() {
                    if matches!(op.unwrap(), wasmparser::Operator::GlobalGet { .. }) {
                        padding_globals += 1;
                    }
                }
            }
            _ => {}
        }
    }
    assert_eq!(signatures.len(), 4);
    for (i, signature) in signatures.iter().enumerate() {
        assert!(!signatures[..i].contains(signature));
    }
    assert_eq!(padding_globals, 0);
}

#[test]
fn arity_specific_imports_remap_calls_across_chunks() {
    let mut chunks = Vec::new();
    for (name, argc) in [("<script>", 2), ("helper", 3)] {
        let mut chunk = Chunk::new(name);
        let push = chunk.add_import("ecma:array", "push");
        for _ in 0..argc {
            chunk.emit_ref_null(HT_EXTERN, 0);
        }
        chunk.emit_call(push, argc, 0);
        chunk.emit_op(Op::DROP, 0);
        chunk.emit_ref_null(HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunks.push(chunk);
    }
    let bytes = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let imports = wasmparser::Parser::new(0)
        .parse_all(&bytes)
        .filter_map(|payload| match payload.unwrap() {
            wasmparser::Payload::ImportSection(imports) => Some(
                imports
                    .into_imports()
                    .filter_map(|import| {
                        let import = import.unwrap();
                        (import.module == "ecma:array" && import.name == "push")
                            .then_some(import.ty)
                    })
                    .collect::<Vec<_>>(),
            ),
            _ => None,
        })
        .next()
        .unwrap();
    assert_eq!(imports.len(), 2);
    assert_ne!(imports[0], imports[1]);
}

#[test]
fn splice_omitted_arguments_keep_typed_prefix_without_padding() {
    let mut chunk = Chunk::new("<script>");
    let splice = chunk.add_import("ecma:array", "splice");
    for argc in [2, 3, 4, 5] {
        chunk.emit_ref_null(HT_EXTERN, 0);
        chunk.emit_i32_const(0, 0);
        if argc >= 3 {
            chunk.emit_i32_const(0, 0);
        }
        for _ in 3..argc {
            chunk.emit_ref_null(HT_EXTERN, 0);
        }
        chunk.emit_call(splice, argc, 0);
        chunk.emit_op(Op::DROP, 0);
    }
    chunk.emit_ref_null(HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .expect("splice must retain its numeric prefix for every observed arity");
    let mut imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                imports.extend(section.into_imports().filter_map(|entry| {
                    let entry = entry.unwrap();
                    (entry.module == "ecma:array" && entry.name == "splice").then_some(entry.ty)
                }));
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                for op in body.get_operators_reader().unwrap() {
                    assert!(
                        !matches!(op.unwrap(), wasmparser::Operator::GlobalGet { .. }),
                        "omitted splice arguments must not be padded with undefined"
                    );
                }
            }
            _ => {}
        }
    }
    assert_eq!(imports.len(), 4);
}

#[test]
fn unshift_accepts_multiple_or_no_items_without_padding() {
    let mut chunk = Chunk::new("<script>");
    let unshift = chunk.add_import("ecma:array", "unshift");
    for argc in [1, 2, 3, 4] {
        for _ in 0..argc {
            chunk.emit_ref_null(HT_EXTERN, 0);
        }
        chunk.emit_call(unshift, argc, 0);
        chunk.emit_op(Op::DROP, 0);
    }
    chunk.emit_ref_null(HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .expect("unshift calls must have their observed Wasm arity");
    let mut imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                imports.extend(section.into_imports().filter_map(|entry| {
                    let entry = entry.unwrap();
                    (entry.module == "ecma:array" && entry.name == "unshift").then_some(entry.ty)
                }));
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                for op in body.get_operators_reader().unwrap() {
                    assert!(
                        !matches!(op.unwrap(), wasmparser::Operator::GlobalGet { .. }),
                        "unshift must not receive padded undefined items"
                    );
                }
            }
            _ => {}
        }
    }
    assert_eq!(imports.len(), 4);
}

#[test]
fn array_constructors_and_optional_mapper_keep_observed_arity() {
    let mut chunk = Chunk::new("<script>");
    for (name, arities) in [
        ("new", &[0, 1, 3][..]),
        ("of", &[0, 1, 3][..]),
        ("from", &[1, 2][..]),
        ("fromAsync", &[1, 2][..]),
    ] {
        let import = chunk.add_import("ecma:array", name);
        for &argc in arities {
            for _ in 0..argc {
                chunk.emit_ref_null(HT_EXTERN, 0);
            }
            chunk.emit_call(import, argc, 0);
            chunk.emit_op(Op::DROP, 0);
        }
    }
    chunk.emit_ref_null(HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .expect("array constructors and mapper forms must validate as Wasm");
    let mut counts = std::collections::HashMap::new();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                for entry in section.into_imports() {
                    let entry = entry.unwrap();
                    if entry.module == "ecma:array" {
                        *counts.entry(entry.name.to_string()).or_insert(0) += 1;
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                for op in body.get_operators_reader().unwrap() {
                    assert!(
                        !matches!(op.unwrap(), wasmparser::Operator::GlobalGet { .. }),
                        "omitted constructor or mapper arguments must not be padded"
                    );
                }
            }
            _ => {}
        }
    }
    assert_eq!(counts.get("new"), Some(&3));
    assert_eq!(counts.get("of"), Some(&3));
    assert_eq!(counts.get("from"), Some(&2));
    assert_eq!(counts.get("fromAsync"), Some(&2));
}
