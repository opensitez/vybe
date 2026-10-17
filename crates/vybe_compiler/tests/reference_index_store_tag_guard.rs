use vybe_ast::{
    BinOp, BindingPattern, BitLane, Directives, EqualityFallback, ExprKind,
    Expression, IntOp, Lang, Module, OperatorDispatch, Param, PassBy, Statement,
    StmtKind, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::{platforms, pointers};
use vybe_runtime::{Chunk, Op, VM, Value};
use vybe_runtime::value::Object;
use vybex as _;

fn member(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object), field: field.into(), null_safe: false,
    })
}

fn module(directives: Directives) -> Module {
    let pointer = Expression::ident("pointer");
    let target = Expression::new(ExprKind::Index {
        object: Box::new(member(pointer.clone(), pointers::CARRAY_BASE_KEY)),
        index: Box::new(member(pointer, pointers::CARRAY_IDX_KEY)),
        null_safe: false,
    });
    let rhs = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("rhs_once")),
        args: vec![], optional: false,
    });
    let function = |name: &str, params, body| Statement::new(StmtKind::FunctionDecl {
        name: name.into(), params, body, return_type: None, modifiers: Default::default(),
        handles: vec![], is_async: false, is_generator: false, is_sub: false,
    });
    Module {
        name: "reference-index-store-tag-guard".into(), language: Lang::Unknown,
        imports: vec![], canon: Default::default(), directives,
        body: vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident("calls".into()), type_hint: Some("int".into()),
                    init: Some(Expression::int(0)), array_bounds: None, with_events: false,
                }], kind: VarDeclKind::Var,
            }),
            function("rhs_once", vec![], vec![
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                    target: Box::new(Expression::ident("calls")),
                    value: Box::new(Expression::new(ExprKind::Binary {
                        op: BinOp::Integer(IntOp::Add, BitLane::W32),
                        left: Box::new(Expression::ident("calls")),
                        right: Box::new(Expression::int(1)),
                    })),
                }))),
                Statement::new(StmtKind::Return(Some(Expression::int(71)))),
            ]),
            function("store", vec![Param {
                name: "pointer".into(), type_hint: None, default: None,
                pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
                is_optional: false, is_nullable: false,
            }], vec![Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Assign {
                target: Box::new(target), value: Box::new(rhs),
            }))))]),
            Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::FuncRef(
                "store".into(),
            ))))),
        ],
    }
}

fn object(fields: &[(&str, Value)]) -> Value {
    let mut object = Object::new();
    for (key, value) in fields {
        object.properties.insert((*key).into(), value.clone());
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn assert_static_guard(chunks: &[Chunk]) {
    for chunk in chunks {
        assert!(!vybe_runtime::debug::disassemble(chunk).contains("dyneq"), "{}", chunk.name);
        let imports = if chunk.imports.is_empty() { &chunks[0].imports } else { &chunk.imports };
        let mut offset = 0;
        while offset < chunk.code.len() {
            let op = Op::from_code_at(&chunk.code, offset);
            if op == Op::CALL {
                let index = chunk.read_u16(offset + 4) as usize;
                if let Some(import) = imports.get(index) {
                    assert!(import.module != "ecma:value"
                        || !matches!(import.name.as_str(), "abstractEq" | "abstractNe"),
                        "{} at {offset}: {}:{}", chunk.name, import.module, import.name);
                }
            }
            offset += 4 + op.operand_format().size_in(&chunk.code, offset + 4);
        }
    }
    let store = chunks.iter().find(|chunk| chunk.name == "store").unwrap();
    let dump = vybe_runtime::debug::disassemble(store);
    for key in ["__base", "__idx", "__ref_kind"] {
        assert_eq!(dump.lines().filter(|line| line.contains("struct.get")
            && line.contains(&format!("({key})"))).count(), 1, "{key}: {dump}");
    }
}

fn wasm_abstract_equality_calls(bytes: &[u8]) -> usize {
    let mut imports = Vec::new();
    let mut calls = 0;
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                for import in section.into_imports() {
                    let import = import.unwrap();
                    if matches!(import.ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_))
                    {
                        imports.push(import.module == "ecma:value"
                            && matches!(import.name, "abstractEq" | "abstractNe"));
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                let mut operators = body.get_operators_reader().unwrap();
                while !operators.eof() {
                    if let wasmparser::Operator::Call { function_index }
                        | wasmparser::Operator::ReturnCall { function_index } = operators.read().unwrap()
                    {
                        calls += usize::from(imports.get(function_index as usize) == Some(&true));
                    }
                }
            }
            _ => {}
        }
    }
    calls
}

#[test]
fn generated_index_store_tag_guard_is_profile_independent_and_noncoercing() {
    for (dispatch, equality, ecma) in [
        (OperatorDispatch::StaticBuiltin, EqualityFallback::Exact, false),
        (OperatorDispatch::Ecma, EqualityFallback::EcmaAbstract, true),
        (OperatorDispatch::RuntimeProtocol, EqualityFallback::Slot, false),
    ] {
        let directives = Directives {
            operator_dispatch: Some(dispatch), equality_fallback: Some(equality),
            ecma_operator_coercion: Some(ecma), ..Default::default()
        };
        let profile = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
        let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&module(directives)).unwrap();
        assert_static_guard(&chunks);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false, include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        assert_eq!(wasm_abstract_equality_calls(&bytes), 0,
            "metadata-free WASM retained actual abstract equality calls");
        let mut vm = VM::new();
        platforms::register_platforms_all(&mut vm);
        let callable = vm.run(chunks).unwrap();
        let mut calls = 0;
        for kind in [
            None, Some(Value::Undefined), Some(Value::Null), Some(Value::I32(1)),
            Some(Value::Bool(true)), Some(object(&[])), Some(Value::String("other".into())),
            Some(Value::String("cell".into())),
        ] {
            let is_cell = matches!(&kind, Some(Value::String(kind)) if kind.as_ref() == "cell");
            let mut fields = vec![("1", Value::I32(9)), ("__value", Value::I32(11))];
            if let Some(kind) = kind { fields.push((pointers::REF_KIND_KEY, kind)); }
            let base = object(&fields);
            let pointer = object(&[
                (pointers::REF_KIND_KEY, Value::String(pointers::CARRAY_KIND.into())),
                (pointers::CARRAY_BASE_KEY, base.clone()),
                (pointers::CARRAY_IDX_KEY, Value::I32(1)),
            ]);
            assert_eq!(vm.invoke(&callable, &[pointer]).unwrap().as_i32(), 71);
            calls += 1;
            assert_eq!(vm.global("calls").unwrap().as_i32(), calls);
            let Value::Object(base) = base else { unreachable!() };
            let base = base.lock().unwrap();
            assert_eq!(base.properties["1"].as_i32(), if is_cell { 9 } else { 71 });
            assert_eq!(base.properties["__value"].as_i32(), if is_cell { 71 } else { 11 });
        }
        let base = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            Value::I32(3), Value::I32(9), Value::I32(5),
        ])));
        let pointer = object(&[
            (pointers::REF_KIND_KEY, Value::String(pointers::CARRAY_KIND.into())),
            (pointers::CARRAY_BASE_KEY, base.clone()),
            (pointers::CARRAY_IDX_KEY, Value::I32(1)),
        ]);
        assert_eq!(vm.invoke(&callable, &[pointer]).unwrap().as_i32(), 71);
        assert_eq!(vm.global("calls").unwrap().as_i32(), calls + 1);
        let Value::Object(base) = base else { unreachable!() };
        let base = base.lock().unwrap();
        let vybe_runtime::value::ObjectKind::Array(values) = &base.kind else { unreachable!() };
        assert_eq!(values.iter().map(Value::as_i32).collect::<Vec<_>>(), vec![3, 71, 5]);
    }
}

#[test]
fn ordinary_ecma_equality_keeps_its_language_fallback() {
    let mut module = module(Directives {
        operator_dispatch: Some(OperatorDispatch::Ecma),
        equality_fallback: Some(EqualityFallback::EcmaAbstract),
        ecma_operator_coercion: Some(true),
        ..Default::default()
    });
    let StmtKind::FunctionDecl { name, params, body, .. } = &mut module.body[2].kind else {
        unreachable!();
    };
    *name = "ordinary_eq".into();
    let mut other = params[0].clone();
    other.name = "other".into();
    params.push(other);
    *body = vec![Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::ident("pointer")),
        right: Box::new(Expression::ident("other")),
    }))))];
    module.body[3] = Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::FuncRef(
        "ordinary_eq".into(),
    )))));
    let profile = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap();
    let ordinary = chunks.iter().find(|chunk| chunk.name == "ordinary_eq").unwrap();
    let imports = if ordinary.imports.is_empty() { &chunks[0].imports } else { &ordinary.imports };
    let mut calls = 0;
    let mut offset = 0;
    while offset < ordinary.code.len() {
        let op = Op::from_code_at(&ordinary.code, offset);
        if op == Op::CALL {
            let index = ordinary.read_u16(offset + 4) as usize;
            calls += usize::from(imports.get(index).is_some_and(|import|
                import.module == "ecma:value" && import.name == "abstractEq"));
        }
        offset += 4 + op.operand_format().size_in(&ordinary.code, offset + 4);
    }
    assert_eq!(calls, 1);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    assert!(wasm_abstract_equality_calls(&bytes) > 0);
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let callable = vm.run(chunks).unwrap();
    assert_eq!(vm.invoke(&callable, &[Value::String("1".into()), Value::I32(1)])
        .unwrap().as_i32(), 1);
    assert_eq!(vm.invoke(&callable, &[Value::String("2".into()), Value::I32(1)])
        .unwrap().as_i32(), 0);
}
