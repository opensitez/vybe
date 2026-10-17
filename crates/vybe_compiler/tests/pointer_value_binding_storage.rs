use vybe_ast::{
    BindingPattern, ExprKind, Expression, Lang, Module, Statement, StmtKind, UnaryOp, VarDeclKind,
    VarDeclarator,
};
use vybex as _;

fn address(name: &str) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr: Box::new(Expression::ident(name)),
    })
}

fn declaration(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.into()),
            type_hint: None,
            init: Some(value),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Var,
    })
}

#[test]
fn address_of_pointer_value_retains_an_independent_binding_cell() {
    let profile =
        vybe_compiler::profile::parse_profile("[compiler]\npromote_addr_taken_at_entry = true\n")
            .unwrap();
    let module = Module {
        name: "pointer-value-storage".into(),
        language: Lang::Unknown,
        imports: vec![],
        canon: Default::default(),
        directives: Default::default(),
        body: vec![
            Statement::new(StmtKind::FunctionDecl {
                name: "pointer_level".into(),
                params: vec![],
                return_type: None,
                body: vec![
                    declaration("value", Expression::int(6)),
                    declaration("pointer", address("value")),
                    declaration("outer", address("pointer")),
                    Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::RefLoad(
                        Box::new(Expression::ident("outer")),
                    ))))),
                ],
                modifiers: Default::default(),
                handles: vec![],
                is_async: false,
                is_generator: false,
                is_sub: false,
            }),
            Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("pointer_level")),
                args: vec![],
                optional: false,
            })))),
        ],
    };
    let chunks = vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .expect("pointer storage must emit standard Wasm");
    let mut vm = vybe_runtime::VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let result = vm.run(chunks).unwrap();
    let vybe_runtime::Value::Object(pointer) = result else {
        panic!("one dereference must return the stored pointer, not its pointee: {result:?}");
    };
    assert_eq!(
        pointer
            .lock()
            .unwrap()
            .properties
            .get("__value")
            .map(|value| value.as_i32()),
        Some(6)
    );
}
