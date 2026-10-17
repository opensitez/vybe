use vybe_runtime::{Chunk, Op, VM};

fn byte_mask(value: &vybe_ast::Expression) -> bool {
    match &value.kind {
        vybe_ast::ExprKind::Lit(vybe_ast::Literal::Int(255)) => true,
        vybe_ast::ExprKind::Cast { expr, type_name }
            if matches!(type_name.as_str(), "int32" | "uint32") =>
        {
            byte_mask(expr)
        }
        _ => false,
    }
}

fn chunks(byte: i32, optimized: bool, floating: bool) -> Vec<Chunk> {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(16, 0);
    if floating {
        chunk.emit_f64_const(byte as f64, 0);
    } else {
        chunk.emit_i32_const(byte, 0);
    }
    let mut chunks = vec![chunk];
    if optimized {
        vybe_compiler::primitives::pointers::emit_linear_byte_store(&mut chunks, 0, 0);
    } else {
        vybe_compiler::primitives::pointers::emit_linear_i32_store8(&mut chunks, 0, 0);
    }
    let chunk = &mut chunks[0];
    chunk.emit_i32_const(byte, 0);
    chunk.emit_op(Op::I32_NE, 0);
    chunk.emit_i32_const(16, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    chunk.emit_i32_const(byte, 0);
    chunk.emit_op(Op::I32_NE, 0);
    chunk.emit_op(Op::I32_OR, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunks
}

#[test]
fn c_byte_stores_preserve_operand_effects_and_wrapping() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/proven_byte_store.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let mut prepared = bundle.prepared_module().unwrap();
    let mut stores = 0;
    let mut masked_narrowings = 0;
    let mut masked_words = 0;
    for statement in &mut prepared.body {
        statement.walk_exprs_mut(&mut |value| {
            if matches!(&value.kind, vybe_ast::ExprKind::Call { callee, .. }
                if matches!(&callee.kind, vybe_ast::ExprKind::Ident(name)
                    if name == "__c_ptr_byte_store"))
            {
                stores += 1;
            }
            if matches!(&value.kind, vybe_ast::ExprKind::Cast { expr, type_name }
                if type_name == "uint8"
                    && matches!(&expr.kind, vybe_ast::ExprKind::Binary {
                        op: vybe_ast::BinOp::Integer(vybe_ast::IntOp::And, vybe_ast::BitLane::W32),
                        left, right,
                    } if byte_mask(left) || byte_mask(right)))
            {
                masked_narrowings += 1;
            }
            if matches!(&value.kind, vybe_ast::ExprKind::Binary {
                op: vybe_ast::BinOp::Integer(vybe_ast::IntOp::And, vybe_ast::BitLane::W32),
                left, right,
            } if byte_mask(left) || byte_mask(right))
            {
                masked_words += 1;
            }
        });
    }
    assert!(stores > 0, "byte-store fast path not selected");
    assert_eq!(masked_narrowings, 0, "masked words are already bytes");
    assert!(
        masked_words >= 2,
        "both mask operand orders must be exercised"
    );
    let chunks = bundle.compile().unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn proven_byte_store_preserves_value_and_removes_two_conversions() {
    vybe_language_c::register();
    for byte in [0, 1, 127, 128, 255] {
        for floating in [false, true] {
            let original = chunks(byte, false, floating);
            let optimized = chunks(byte, true, floating);
            let count = |chunks: &[Chunk]| {
                vybe_runtime::debug::disassemble(&chunks[0])
                    .lines()
                    .filter(|line| line.as_bytes().first().is_some_and(u8::is_ascii_digit))
                    .count()
            };
            assert_eq!(count(&original) - count(&optimized), 2);
            for chunks in [original, optimized] {
                let bytes = vybe_platform_wasm::write_wasm_with_options(
                    &chunks,
                    vybe_platform_wasm::WasmWriteOptions {
                        include_vybe_metadata: false,
                        include_name_section: true,
                    },
                );
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&bytes)
                    .unwrap();
                assert_eq!(VM::new().run(chunks).unwrap().as_i32(), 0);
                let reloaded = vybe_platform_wasm::read_wasm(&bytes).unwrap();
                let mut vm = VM::new();
                vybe_runtime::init_all_registered(
                    &mut vm,
                    &vybe_runtime::capabilities::Capabilities::all(),
                );
                vm.run(reloaded).unwrap();
                let mut function = vybe_runtime::value::Object::new();
                function.kind =
                    vybe_runtime::value::ObjectKind::Function(vybe_runtime::value::Function {
                        name: None,
                        arity: 0,
                        chunk_index: 1,
                        upvalues: Vec::new(),
                    });
                let function = vybe_runtime::Value::Object(std::sync::Arc::new(
                    std::sync::Mutex::new(function),
                ));
                assert_eq!(vm.invoke(&function, &[]).unwrap().as_i32(), 0);
            }
        }
    }
}
