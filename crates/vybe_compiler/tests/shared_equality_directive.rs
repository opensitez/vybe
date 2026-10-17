use vybe_ast::{
    BinOp, BindingPattern, Directives, EqualityFallback, ExprKind, Expression, Lang, Module,
    OperatorDispatch, Statement, StmtKind, TypeResolution, VarDeclKind, VarDeclarator,
};
use vybe_runtime::VM;

fn evaluate(left: Expression, right: Expression, policy: EqualityFallback) -> bool {
    evaluate_with_materialization(left, right, policy, true)
}

fn evaluate_with_materialization(
    left: Expression, right: Expression, policy: EqualityFallback, boxed: bool,
) -> bool {
    let declarations = [("left", left), ("right", right)]
        .into_iter()
        .map(|(name, init)| {
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name.into()),
                    type_hint: None,
                    init: Some(init),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Var,
            })
        });
    let mut body: Vec<_> = declarations.collect();
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::Ident("left".into()))),
            right: Box::new(Expression::new(ExprKind::Ident("right".into()))),
        },
    )))));
    let module = Module {
        name: "shared-equality-directive".into(),
        language: Lang::Unknown,
        body,
        imports: vec![],
        canon: Default::default(),
        directives: Directives {
            equality_fallback: Some(policy),
            operator_dispatch: Some(OperatorDispatch::StaticBuiltin),
            type_resolution: Some(TypeResolution::Dynamic),
            ..Default::default()
        },
    };
    let profile = vybe_compiler::profile::parse_profile(&format!(
        "[compiler]\nmaterialize_bool_results = {boxed}\n",
    ))
    .unwrap();
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
        .unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let result = vm.run(chunks).unwrap();
    result.as_f64() != 0.0
}

#[test]
fn numeric_boolean_representation_retains_numeric_equality() {
    assert!(evaluate_with_materialization(
        Expression::bool(true), Expression::int(1), EqualityFallback::Exact, false,
    ));
    assert!(!evaluate_with_materialization(
        Expression::bool(false), Expression::int(1), EqualityFallback::Exact, false,
    ));
}

#[test]
fn integer_logical_policy_inherits_and_explicit_false_overrides() {
    let mut region = Directives {
        logical_ops_bitwise_for_integers: Some(true),
        ..Default::default()
    };
    region.overlay(&Directives::default());
    assert_eq!(region.logical_ops_bitwise_for_integers, Some(true));
    region.overlay(&Directives {
        logical_ops_bitwise_for_integers: Some(false),
        ..Default::default()
    });
    assert_eq!(region.logical_ops_bitwise_for_integers, Some(false));
}

#[test]
fn exact_equality_retains_primitive_and_null_rules() {
    for (left, right, expected) in [
        (Expression::int(3), Expression::float(3.0), true),
        (Expression::int(3), Expression::int(4), false),
        (Expression::float(-0.0), Expression::float(0.0), true),
        (Expression::float(f64::NAN), Expression::float(f64::NAN), false),
        (Expression::string("same"), Expression::string("same"), true),
        (Expression::string("same"), Expression::string("other"), false),
        (Expression::bool(true), Expression::bool(true), true),
        (Expression::bool(true), Expression::bool(false), false),
        (Expression::null(), Expression::null(), true),
        (Expression::null(), Expression::int(0), false),
        (Expression::string("3"), Expression::int(3), false),
        (Expression::bool(true), Expression::int(1), false),
    ] {
        assert_eq!(evaluate(left, right, EqualityFallback::Exact), expected);
    }
}

#[test]
fn ecma_coercion_is_selected_by_directive_not_language_identity() {
    assert!(evaluate(
        Expression::string("3"), Expression::int(3), EqualityFallback::EcmaAbstract,
    ));
    assert!(evaluate(
        Expression::bool(true), Expression::int(1), EqualityFallback::EcmaAbstract,
    ));
}

#[test]
fn integer_slash_policy_inherits_and_explicit_false_overrides() {
    let mut region = Directives {
        integer_division_on_slash: Some(true),
        ..Default::default()
    };
    region.overlay(&Directives::default());
    assert_eq!(region.integer_division_on_slash, Some(true));
    region.overlay(&Directives {
        integer_division_on_slash: Some(false),
        ..Default::default()
    });
    assert_eq!(region.integer_division_on_slash, Some(false));
}

#[test]
fn xor_and_integer_cast_policies_inherit_and_override() {
    let mut region = Directives {
        xor_is_logical_for_non_integers: Some(true),
        integer_cast_truncates: Some(true),
        ..Default::default()
    };
    region.overlay(&Directives::default());
    assert_eq!(region.xor_is_logical_for_non_integers, Some(true));
    assert_eq!(region.integer_cast_truncates, Some(true));
    region.overlay(&Directives {
        xor_is_logical_for_non_integers: Some(false),
        integer_cast_truncates: Some(false),
        ..Default::default()
    });
    assert_eq!(region.xor_is_logical_for_non_integers, Some(false));
    assert_eq!(region.integer_cast_truncates, Some(false));
}

#[test]
fn concatenation_policy_inherits_and_explicit_false_overrides() {
    let mut region = Directives {
        concat_stringifies_operands: Some(true),
        ..Default::default()
    };
    region.overlay(&Directives::default());
    assert_eq!(region.concat_stringifies_operands, Some(true));
    region.overlay(&Directives {
        concat_stringifies_operands: Some(false),
        ..Default::default()
    });
    assert_eq!(region.concat_stringifies_operands, Some(false));
}

#[test]
fn dynamic_numeric_policy_inherits_and_explicit_false_overrides() {
    let mut region = Directives {
        dynamic_numeric_dispatch: Some(true),
        ..Default::default()
    };
    region.overlay(&Directives::default());
    assert_eq!(region.dynamic_numeric_dispatch, Some(true));
    region.overlay(&Directives {
        dynamic_numeric_dispatch: Some(false),
        ..Default::default()
    });
    assert_eq!(region.dynamic_numeric_dispatch, Some(false));
}

#[test]
fn operator_coercion_inherits_and_explicit_false_overrides() {
    let mut region = vybe_ast::Directives {
        ecma_operator_coercion: Some(true),
        ..Default::default()
    };
    region.overlay(&vybe_ast::Directives::default());
    assert_eq!(region.ecma_operator_coercion, Some(true));
    region.overlay(&vybe_ast::Directives {
        ecma_operator_coercion: Some(false),
        ..Default::default()
    });
    assert_eq!(region.ecma_operator_coercion, Some(false));
}

#[test]
fn string_relational_policy_inherits_and_explicit_false_overrides() {
    let mut region = Directives {
        string_aware_relational: Some(true),
        ..Default::default()
    };
    region.overlay(&Directives::default());
    assert_eq!(region.string_aware_relational, Some(true));
    region.overlay(&Directives {
        string_aware_relational: Some(false),
        ..Default::default()
    });
    assert_eq!(region.string_aware_relational, Some(false));
}

#[test]
fn string_addition_policy_inherits_and_explicit_false_overrides() {
    let mut region = Directives {
        add_concatenates_strings: Some(true),
        ..Default::default()
    };
    region.overlay(&Directives::default());
    assert_eq!(region.add_concatenates_strings, Some(true));
    region.overlay(&Directives {
        add_concatenates_strings: Some(false),
        ..Default::default()
    });
    assert_eq!(region.add_concatenates_strings, Some(false));
}
