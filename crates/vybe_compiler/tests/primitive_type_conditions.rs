use vybe_ast::{ExprKind, Expression, Lang, Module, Statement, StmtKind};
use vybe_compiler::primitives::platforms;
use vybe_runtime::{Value, VM};

#[test]
fn primitive_type_tests_keep_conditions_native_and_values_boolean() {
    let profile = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
    for (type_name, operand, expected) in [
        ("number", Expression::int(42), true),
        ("number", Expression::string("42"), false),
        ("string", Expression::string("hello"), true),
        ("string", Expression::null(), false),
        ("boolean", Expression::bool(false), true),
        ("boolean", Expression::int(1), false),
        ("array", Expression::new(ExprKind::Array(vec![])), true),
        ("array", Expression::int(1), false),
        ("object", Expression::new(ExprKind::Object(vec![])), true),
        ("object", Expression::null(), false),
        ("undefined", Expression::null(), false),
        ("function", Expression::int(1), false),
        (
            "function",
            Expression::new(ExprKind::Lambda {
                params: vec![],
                body: vybe_ast::LambdaBody::Expr(Box::new(Expression::null())),
                is_async: false,
                captures: vec![],
            }),
            true,
        ),
        ("bigint", Expression::int(1), false),
        ("symbol", Expression::int(1), false),
    ] {
        for condition in [false, true] {
            let predicate = Expression::new(ExprKind::IsType {
                expr: Box::new(operand.clone()),
                type_name: type_name.into(),
            });
            let result = if condition {
                Expression::new(ExprKind::Ternary {
                    cond: Box::new(predicate),
                    then: Box::new(Expression::int(11)),
                    else_: Box::new(Expression::int(22)),
                })
            } else {
                predicate
            };
            let module = Module {
                name: "<primitive-type-test>".into(),
                language: Lang::Unknown,
                body: vec![Statement::new(StmtKind::Return(Some(result)))],
                imports: vec![],
                canon: Default::default(),
                directives: Default::default(),
            };
            let chunks = vybe_compiler::Compiler::with_profile(profile.clone())
                .compile(&module)
                .unwrap();
            let entry = &chunks[0];
            let dump = vybe_runtime::debug::disassemble(entry);
            if condition {
                assert!(!dump.contains("__vybe_autoderef"), "{type_name}: {dump}");
            }
            let bytes = vybe_platform_wasm::write_wasm(&chunks);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap_or_else(|error| {
                    panic!("{type_name} condition={condition}: {error}\n{dump}")
                });
            let mut vm = VM::new();
            platforms::register_platforms_all(&mut vm);
            let result = vm.run(chunks).unwrap();
            if condition {
                assert_eq!(
                    result.as_i32(),
                    if expected { 11 } else { 22 },
                    "{type_name}"
                );
            } else {
                assert_eq!(result, Value::Bool(expected), "{type_name}");
            }
        }
    }
}
