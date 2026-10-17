use vybe_runtime::{Chunk, Op};
use vybe_runtime::opcode::heaptype::{HeapType, HT_EXTERN, HT_FUNC};

#[test]
fn guest_tables_keep_types_limits_and_indices_separate_from_callables() {
    let mut chunk = Chunk::new("declared-tables");
    chunk.table_min_sizes = vec![1, 41, 2];
    chunk.table_max_sizes = vec![Some(2), Some(41), Some(3)];
    chunk.table_is_64 = vec![false, false, true];
    chunk.table_element_types = vec![
        (HeapType::Abstract(HT_FUNC), true, false, false),
        (HeapType::Abstract(HT_EXTERN), true, false, false),
        (HeapType::Abstract(HT_EXTERN), true, false, false),
    ];
    chunk.emit_ref_null(HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);
    let bytes = vybe_platform_wasm::write_wasm(&[chunk.clone()]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let mut tables = Vec::new();
    let mut populated = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::TableSection(section) => {
                for table in section { tables.push(table.unwrap().ty); }
            }
            wasmparser::Payload::ElementSection(section) => {
                for segment in section {
                    if let wasmparser::ElementKind::Active { table_index, .. } = segment.unwrap().kind {
                        populated.push(table_index.unwrap_or(0));
                    }
                }
            }
            _ => {}
        }
    }
    assert_eq!(tables.len(), 4);
    assert_eq!(tables[0].element_type, wasmparser::RefType::FUNCREF);
    assert_eq!(tables[1].element_type, wasmparser::RefType::EXTERNREF);
    assert_eq!(tables[2].element_type, wasmparser::RefType::EXTERNREF);
    assert_eq!(tables.iter().map(|t| (t.initial, t.maximum, t.table64)).collect::<Vec<_>>(),
        [(1, Some(2), false), (41, Some(41), false), (2, Some(3), true), (1, None, false)]);
    assert_eq!(populated, [3]);
    let decoded = vybe_platform_wasm::read_wasm(&bytes).unwrap();
    assert_eq!(decoded[0].table_min_sizes, [1, 41, 2, 1]);
    assert_eq!(decoded[0].table_max_sizes, [Some(2), Some(41), Some(3), None]);
    assert_eq!(decoded[0].table_element_types[1], chunk.table_element_types[1]);
    let options = vybe_platform_wasm::writer::WasmWriteOptions {
        include_vybe_metadata: true,
        ..Default::default()
    };
    let cached = vybe_platform_wasm::writer::write_wasm_with_options(&[chunk.clone()], options);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&cached).unwrap();
    let decoded = vybe_platform_wasm::read_wasm(&cached).unwrap();
    assert_eq!(decoded[0].table_min_sizes, chunk.table_min_sizes);
    assert_eq!(decoded[0].table_max_sizes, chunk.table_max_sizes);
    assert_eq!(decoded[0].table_element_types, chunk.table_element_types);
    assert_eq!(decoded[0].table_is_64, chunk.table_is_64);
    assert_eq!(vybe_platform_wasm::writer::write_wasm_with_options(&decoded, options), cached);
}

#[test]
fn declared_table_widths_keep_core_operands_native() {
    let mut get32 = Chunk::new("get32");
    get32.func_sig = Some((String::new(), "i32".into()));
    get32.table_min_sizes = vec![2, 2];
    get32.table_max_sizes = vec![Some(4), Some(4)];
    get32.table_is_64 = vec![false, true];
    get32.table_element_types = vec![(HeapType::Abstract(HT_EXTERN), true, false, false); 2];
    get32.emit_i32_const(1, 0);
    get32.emit_op_idx(Op::TABLE_GET, 0u32, 0);
    get32.emit_op(Op::REF_IS_NULL, 0);
    get32.emit_op(Op::RETURN, 0);

    let mut get64 = Chunk::new("get64");
    get64.func_sig = Some((String::new(), "i32".into()));
    get64.emit_i64_const(1, 0);
    get64.emit_op_idx(Op::TABLE_GET, 1u32, 0);
    get64.emit_op(Op::REF_IS_NULL, 0);
    get64.emit_op(Op::RETURN, 0);

    let mut set64 = Chunk::new("set64");
    set64.func_sig = Some((String::new(), "i32".into()));
    set64.emit_i64_const(1, 0);
    set64.emit_ref_null(HT_EXTERN, 0);
    set64.emit_op_idx(Op::TABLE_SET, 1u32, 0);
    set64.emit_i64_const(1, 0);
    set64.emit_op_idx(Op::TABLE_GET, 1u32, 0);
    set64.emit_op(Op::REF_IS_NULL, 0);
    set64.emit_op(Op::RETURN, 0);

    let mut grow64 = Chunk::new("grow64");
    grow64.func_sig = Some((String::new(), "i64".into()));
    grow64.emit_ref_null(HT_EXTERN, 0);
    grow64.emit_i64_const(1, 0);
    grow64.emit_op_idx(Op::TABLE_GROW, 1u32, 0);
    grow64.emit_op(Op::RETURN, 0);
    let bytes = vybe_platform_wasm::write_wasm(&[get32, get64, set64, grow64]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let mut checked = 0;
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
            if checked == 4 { break; }
            for operator in body.get_operators_reader().unwrap() {
                assert!(!matches!(operator.unwrap(), wasmparser::Operator::Call { .. }),
                    "native table operands must not cross boxed helper calls");
            }
            checked += 1;
        }
    }
    assert_eq!(checked, 4);
}

#[test]
#[ignore = "requires a CLI-emitted source artifact with roundtrip metadata"]
fn source_table_artifact_reemits_valid_standard_wasm() {
    let path = std::env::var("VYBE_TABLE_SOURCE_ARTIFACT")
        .expect("set VYBE_TABLE_SOURCE_ARTIFACT to the prepared source artifact");
    let prepared = std::fs::read(path).unwrap();
    let chunks = vybe_platform_wasm::read_wasm(&prepared).unwrap();
    assert!(!chunks[0].table_element_types.is_empty());
    let binary = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary).unwrap();
}
