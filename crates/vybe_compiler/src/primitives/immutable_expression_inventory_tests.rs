use super::*;

fn function(name: &str, body: Vec<Statement>) -> Statement {
    Statement::new(StmtKind::FunctionDecl {
        name: name.into(),
        params: vec![],
        return_type: None,
        body,
        modifiers: Default::default(),
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}
fn root(expression: Expression) -> Statement {
    Statement::new(StmtKind::Expr(expression))
}
fn function_expr(name: &str, body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::FunctionExpr(Box::new(function(name, body))))
}
fn lambda(body: LambdaBody) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: vec![],
        body,
        captures: vec![],
        is_async: false,
    })
}
fn class(name: Option<&str>, members: Vec<ClassMember>) -> Expression {
    Expression::new(ExprKind::ClassExpr {
        name: name.map(str::to_owned),
        parent: None,
        interfaces: vec![],
        members,
    })
}
fn aliases(name: &str) -> Vec<String> {
    vec![
        format!("__vybe_func${name}"),
        format!("__php_fn_{}", name.replace('.', "__")),
    ]
}
fn inventory(expression: Expression) -> Bindings {
    Bindings::prepare_with_aliases(
        &[function("target", vec![]), root(expression)],
        str::to_owned,
        aliases,
    )
}

#[test]
fn unrelated_named_function_expression_retains_only_source_owned_candidate() {
    let source = vec![
        function("target", vec![]),
        function("startup", vec![root(function_expr("guest", vec![]))]),
    ];
    let bindings = Bindings::prepare_with_aliases(&source, str::to_owned, aliases);
    assert_eq!(bindings.by_binding.get("target"), Some(&0));
    assert_eq!(bindings.by_binding.get("startup"), Some(&1));
    assert_eq!(bindings.claimed.get("guest"), Some(&None));
    assert_eq!(bindings.claimed.get("__vybe_func$guest"), Some(&None));
    assert!(!bindings.by_binding.contains_key("guest"));
}

#[test]
fn function_expression_own_name_collision_tombstones_only_that_binding() {
    let bindings = inventory(function_expr("target", vec![]));
    assert!(!bindings.by_binding.contains_key("target"));
    assert_eq!(bindings.claimed.get("__vybe_func$target"), Some(&None));
}

#[test]
fn lambda_expression_and_block_retain_unrelated_candidate() {
    for expression in [
        lambda(LambdaBody::Expr(Box::new(Expression::int(1)))),
        lambda(LambdaBody::Block(vec![Statement::new(StmtKind::Return(
            None,
        ))])),
    ] {
        assert_eq!(inventory(expression).by_binding.get("target"), Some(&0));
    }
}

#[test]
fn conditional_nested_publication_inside_lambda_is_never_owned() {
    let nested = Statement::new(StmtKind::If {
        cond: Expression::bool(false),
        then_body: vec![function("target", vec![])],
        elifs: vec![],
        else_body: None,
    });
    let bindings = inventory(lambda(LambdaBody::Block(vec![nested])));
    assert!(!bindings.by_binding.contains_key("target"));
    assert_eq!(bindings.claimed.get("__vybe_func$target"), Some(&None));
}

#[test]
fn alias_collisions_inside_expression_bodies_honor_exact_and_folded_spelling() {
    for fold in [false, true] {
        let source = vec![
            function("Outer.Inner", vec![]),
            root(lambda(LambdaBody::Block(vec![function(
                if fold { "OUTER__INNER" } else { "Outer__Inner" },
                vec![],
            )]))),
        ];
        let bindings = Bindings::prepare_with_aliases(
            &source,
            |name| {
                if fold {
                    name.to_ascii_lowercase()
                } else {
                    name.into()
                }
            },
            aliases,
        );
        let alias = if fold {
            "__php_fn_outer__inner"
        } else {
            "__php_fn_Outer__Inner"
        };
        assert!(bindings.excluded.contains(alias));
        assert_eq!(bindings.claimed.get(alias), Some(&None));
        assert!(
            bindings
                .by_binding
                .contains_key(if fold { "outer.inner" } else { "Outer.Inner" })
        );
    }
}

#[test]
fn named_class_expression_method_is_not_a_function_global_but_nested_body_is() {
    let expression = class(
        Some("Container"),
        vec![ClassMember::Method(Box::new(function(
            "target",
            vec![function("nested", vec![])],
        )))],
    );
    let bindings = inventory(expression);
    assert_eq!(bindings.by_binding.get("target"), Some(&0));
    assert_eq!(bindings.claimed.get("Container"), Some(&None));
    assert_eq!(bindings.claimed.get("nested"), Some(&None));
    assert_eq!(bindings.claimed.get("__vybe_func$target"), Some(&Some(0)));
}

#[test]
fn class_expression_own_binding_and_field_initializer_publications_are_seen() {
    assert!(
        !inventory(class(Some("target"), vec![]))
            .by_binding
            .contains_key("target")
    );
    let field = ClassMember::Field {
        name: "field".into(),
        type_hint: None,
        init: Some(function_expr("target", vec![])),
        modifiers: Default::default(),
        with_events: false,
        array_bounds: None,
        storage: None,
    };
    assert!(
        !inventory(class(Some("Container"), vec![field]))
            .by_binding
            .contains_key("target")
    );
}

#[test]
fn class_method_defaults_and_constructor_arguments_are_not_hidden_surfaces() {
    let mut method = function("method", vec![]);
    let StmtKind::FunctionDecl { params, .. } = &mut method.kind else {
        unreachable!()
    };
    params.push(Param {
        name: "argument".into(),
        type_hint: None,
        default: Some(function_expr("target", vec![])),
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: false,
    });
    assert!(
        !inventory(class(
            Some("Container"),
            vec![ClassMember::Method(Box::new(method))]
        ))
        .by_binding
        .contains_key("target")
    );
    let constructor = ClassMember::Constructor {
        name: None,
        params: vec![],
        body: vec![],
        base_args: Some(vec![function_expr("target", vec![])]),
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    };
    assert!(
        !inventory(class(Some("Container"), vec![constructor]))
            .by_binding
            .contains_key("target")
    );
}

#[test]
fn deeply_nested_expression_publications_do_not_get_source_ownership() {
    let nested = lambda(LambdaBody::Expr(Box::new(function_expr(
        "guest",
        vec![root(class(
            Some("Container"),
            vec![ClassMember::Method(Box::new(function(
                "method",
                vec![function("target", vec![])],
            )))],
        ))],
    ))));
    let bindings = inventory(nested);
    assert!(!bindings.by_binding.contains_key("target"));
    assert!(!bindings.by_binding.contains_key("guest"));
}

#[test]
fn writes_and_explicit_reference_captures_inside_expressions_exclude_candidates() {
    let write = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident("target")],
        value: Expression::int(0),
        by_ref: false,
    });
    assert!(
        !inventory(lambda(LambdaBody::Block(vec![write])))
            .by_binding
            .contains_key("target")
    );
    let expression = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Expr(Box::new(Expression::int(0))),
        captures: vec!["&target".into()],
        is_async: false,
    });
    assert!(!inventory(expression).by_binding.contains_key("target"));
}

#[test]
fn anonymous_or_state_dependent_expression_publishers_keep_whole_inventory_rejection() {
    let mut parent = class(Some("Container"), vec![]);
    let ExprKind::ClassExpr { parent: slot, .. } = &mut parent.kind else {
        unreachable!()
    };
    *slot = Some(Box::new(Expression::ident("Base")));
    for expression in [function_expr("", vec![]), class(None, vec![]), parent] {
        assert!(inventory(expression).by_binding.is_empty());
    }
}

#[test]
fn unknown_statement_surface_inside_expression_keeps_whole_inventory_rejection() {
    let unknown = Statement::new(StmtKind::NamespaceDecl {
        name: "opaque".into(),
        body: vec![],
    });
    for expression in [
        function_expr("guest", vec![unknown.clone()]),
        lambda(LambdaBody::Block(vec![unknown])),
    ] {
        assert!(inventory(expression).by_binding.is_empty());
    }
}

#[test]
fn generated_exit_status_is_a_known_recursively_inventoried_expression_container() {
    let exit = Statement::new(StmtKind::Exit { status: Some(Expression::int(0)) });
    assert_eq!(inventory(function_expr("guest", vec![exit])).by_binding.get("target"), Some(&0));
    let exit = Statement::new(StmtKind::Exit { status: None });
    assert_eq!(inventory(function_expr("guest", vec![exit])).by_binding.get("target"), Some(&0));
    let exit = Statement::new(StmtKind::Exit { status: Some(function_expr("target", vec![])) });
    assert!(!inventory(function_expr("guest", vec![exit])).by_binding.contains_key("target"));
    let write = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident("target")), value: Box::new(Expression::int(0)),
    });
    let exit = Statement::new(StmtKind::Exit { status: Some(write) });
    assert!(!inventory(function_expr("guest", vec![exit])).by_binding.contains_key("target"));
    let unknown = Statement::new(StmtKind::NamespaceDecl { name: "opaque".into(), body: vec![] });
    let exit = Statement::new(StmtKind::Exit { status: Some(function_expr("guest2", vec![unknown])) });
    assert!(inventory(function_expr("guest", vec![exit])).by_binding.is_empty());
}

#[test]
fn nested_statement_traversal_cost_counts_raw_callbacks_not_only_deduplicated_facts() {
    fn old_recursive_cost(statement: &Statement, count: &mut usize) {
        statement.walk_exprs(&mut |_| *count += 1);
        if let StmtKind::Block(body) = &statement.kind {
            for statement in body { old_recursive_cost(statement, count); }
        }
    }
    for depth in [32, 64, 128, 256] {
        let mut statement = root(Expression::int(0));
        for _ in 0..depth {
            statement = Statement::new(StmtKind::Block(vec![root(Expression::int(0)), statement]));
        }
        let address = &statement as *const Statement as usize;
        let mut old = 0;
        old_recursive_cost(&statement, &mut old);
        let mut walk = BindingWalk::default();
        let mut statements = 0;
        let mut expressions = 0;
        assert!(visit_binding_tree(&statement, &mut walk, &mut |node| match node {
            BindingNode::Statement(_) => statements += 1,
            BindingNode::Expression(_) => expressions += 1,
        }));
        assert_eq!(walk.expression_callbacks, depth + 1);
        assert_eq!(expressions, depth + 1);
        assert_eq!(statements, depth * 2 + 1);
        assert_eq!(old, depth * (depth + 1) / 2 + 2 * depth + 1);
        assert_eq!(&statement as *const Statement as usize, address);
        eprintln!("nested-block depth={depth} old_recursive_callbacks={old} new_raw_callbacks={} statements={statements}", walk.expression_callbacks);
    }
}

#[test]
fn nested_function_expression_traversal_visits_each_expression_once() {
    for depth in [16, 32, 64, 128] {
        let mut expression = Expression::int(0);
        for index in 0..depth {
            expression = function_expr(&format!("guest{index}"), vec![root(expression)]);
        }
        let source = root(expression);
        let mut walk = BindingWalk::default();
        assert!(visit_binding_tree(&source, &mut walk, &mut |_| {}));
        assert_eq!(walk.expression_callbacks, depth + 1);
        assert_eq!(walk.expressions.len(), depth + 1);
        assert_eq!(walk.statements.len(), depth * 2 + 1);
        eprintln!("nested-function-expression depth={depth} raw_callbacks={} statements={}", walk.expression_callbacks, walk.statements.len());
    }
}

#[test]
fn exit_status_does_not_rewalk_an_expression_root_already_seen() {
    let status = function_expr("guest", vec![root(Expression::int(0))]);
    let source = Statement::new(StmtKind::Exit { status: Some(status) });
    let mut walk = BindingWalk::default();
    assert!(visit_binding_tree(&source, &mut walk, &mut |_| {}));
    assert_eq!(walk.expression_callbacks, 2);
    let StmtKind::Exit { status: Some(status) } = &source.kind else { unreachable!() };
    assert!(visit_binding_expression(status, &mut walk, &mut |_| {}));
    assert_eq!(walk.expression_callbacks, 2);
}

#[test]
fn assertion_test_and_message_are_both_recursively_inventoried() {
    let assertion = |test, msg| Statement::new(StmtKind::Assert { test, msg });
    assert_eq!(inventory(function_expr("guest", vec![assertion(Expression::bool(true), None)]))
        .by_binding.get("target"), Some(&0));
    for (test, msg) in [(function_expr("target", vec![]), None),
        (Expression::bool(true), Some(function_expr("target", vec![])))] {
        assert!(!inventory(function_expr("guest", vec![assertion(test, msg)])).by_binding.contains_key("target"));
    }
}

#[test]
fn data_segment_offsets_are_conservatively_inventoried_even_before_constant_validation() {
    let segment = |offset| Statement::new(StmtKind::DataSegment {
        memory_index: 0, offset, bytes: vec![1, 2, 3],
    });
    for offset in [None, Some(Expression::int(32))] {
        assert_eq!(inventory(function_expr("guest", vec![segment(offset)])).by_binding.get("target"), Some(&0));
    }
    // The emitter later rejects this nonconstant offset; the inventory must
    // not grant a binding receipt by ignoring its publication surface first.
    assert!(!inventory(function_expr("guest", vec![segment(Some(function_expr("target", vec![])))]))
        .by_binding.contains_key("target"));
}

#[test]
fn numeric_memory_and_label_only_goto_have_no_named_binding_publication() {
    let source = vec![function("target", vec![]),
        Statement::new(StmtKind::MemoryDecl { min_pages: 1, max_pages: Some(2), is_64: false }),
        Statement::new(StmtKind::GoTo("target".into()))];
    let bindings = Bindings::prepare(&source, str::to_owned);
    assert_eq!(bindings.by_binding.get("target"), Some(&0));
    assert!(!bindings.excluded.contains("target"));
}

#[test]
fn only_examined_numeric_transport_directive_deltas_keep_candidate_inventory() {
    for scope in [DirectiveScope::Block, DirectiveScope::Module] {
        for transport in [None, Some(WideIntegerTransport::Exact), Some(WideIntegerTransport::Number)] {
            let source = vec![function("target", vec![]), Statement::new(StmtKind::Directive {
                set: Directives { wide_integer_transport: transport, ..Default::default() }, scope,
            })];
            assert_eq!(Bindings::prepare(&source, str::to_owned).by_binding.get("target"), Some(&0));
        }
    }
}

#[test]
fn directive_binding_canonicalization_and_unexamined_fields_still_reject_inventory() {
    for set in [Directives { variable_case: Some(CaseMatch::Folded), ..Default::default() },
        Directives { callable_case: Some(CaseMatch::Exact), ..Default::default() },
        Directives { declared_function_bindings_live: Some(false), ..Default::default() },
        Directives { receiver_binding: Some(ReceiverBinding::UniversalParameter), ..Default::default() },
        Directives { wide_integer_transport: Some(WideIntegerTransport::Exact),
            operator_dispatch: Some(OperatorDispatch::StaticBuiltin), ..Default::default() }] {
        let source = vec![function("target", vec![]), Statement::new(StmtKind::Directive {
            set, scope: DirectiveScope::Module,
        })];
        assert!(Bindings::prepare(&source, str::to_owned).by_binding.is_empty());
    }
}

#[test]
fn unknown_containers_in_assertion_or_segment_expression_still_reject_whole_inventory() {
    let unknown = || function_expr("guest", vec![Statement::new(StmtKind::NamespaceDecl {
        name: "opaque".into(), body: vec![],
    })]);
    for statement in [Statement::new(StmtKind::Assert { test: unknown(), msg: None }),
        Statement::new(StmtKind::DataSegment { memory_index: 0, offset: Some(unknown()), bytes: vec![] })] {
        let source = vec![function("target", vec![]), statement];
        assert!(Bindings::prepare(&source, str::to_owned).by_binding.is_empty());
    }
}

fn object_member(kind: u8, declaration: Statement) -> Expression {
    let property = match kind {
        0 => ObjectProperty::Method { key: "target".into(), value: Box::new(declaration) },
        1 | 2 => ObjectProperty::Accessor {
            kind: if kind == 1 { AccessorKind::Get } else { AccessorKind::Set },
            key: "target".into(), value: Box::new(declaration),
        },
        _ => unreachable!(),
    };
    Expression::new(ExprKind::Object(vec![property]))
}

#[test]
fn object_method_and_accessor_names_are_properties_not_function_globals() {
    for kind in 0..3 {
        let object = object_member(kind, function("target", vec![Statement::new(StmtKind::Return(None))]));
        let bindings = inventory(lambda(LambdaBody::Block(vec![root(object)])));
        assert_eq!(bindings.by_binding.get("target"), Some(&0));
        assert_eq!(bindings.claimed.get("__vybe_func$target"), Some(&Some(0)));
        assert!(!bindings.excluded.contains("target"));
    }
}

#[test]
fn object_member_nested_conditional_alias_publications_are_unknown_owners() {
    for kind in 0..3 {
        let conditional = Statement::new(StmtKind::If {
            cond: Expression::bool(false), then_body: vec![function("Outer__Inner", vec![])],
            elifs: vec![], else_body: None,
        });
        let source = vec![function("Outer.Inner", vec![]),
            root(lambda(LambdaBody::Block(vec![root(object_member(kind,
                function("property_only", vec![conditional])))])))];
        let bindings = Bindings::prepare_with_aliases(&source, str::to_owned, aliases);
        assert_eq!(bindings.by_binding.get("Outer.Inner"), Some(&0));
        assert_eq!(bindings.claimed.get("__php_fn_Outer__Inner"), Some(&None));
        assert!(!bindings.by_binding.contains_key("Outer__Inner"));
        assert!(!bindings.claimed.contains_key("property_only"));
    }
}

#[test]
fn object_member_body_writes_exclude_root_function_bindings() {
    for kind in 0..3 {
        let write = Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident("target")], value: Expression::int(0), by_ref: false,
        });
        let bindings = inventory(object_member(kind, function("property_only", vec![write])));
        assert!(!bindings.by_binding.contains_key("target"));
    }
}

#[test]
fn object_member_parameter_defaults_inventory_publications_and_writes() {
    for kind in 0..3 {
        for default in [function_expr("target", vec![]), Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident("target")), value: Box::new(Expression::int(0)),
        })] {
            let mut declaration = function("property_only", vec![]);
            let StmtKind::FunctionDecl { params, .. } = &mut declaration.kind else { unreachable!() };
            params.push(Param { name: "argument".into(), type_hint: None, default: Some(default),
                pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_optional: true, is_nullable: false });
            assert!(!inventory(object_member(kind, declaration)).by_binding.contains_key("target"));
        }
    }
}

#[test]
fn object_member_unknown_body_and_default_surfaces_reject_the_inventory() {
    for kind in 0..3 {
        let unknown = || Statement::new(StmtKind::NamespaceDecl { name: "opaque".into(), body: vec![] });
        assert!(inventory(object_member(kind, function("property_only", vec![unknown()])))
            .by_binding.is_empty());
        let mut declaration = function("property_only", vec![]);
        let StmtKind::FunctionDecl { params, .. } = &mut declaration.kind else { unreachable!() };
        params.push(Param { name: "argument".into(), type_hint: None,
            default: Some(function_expr("guest", vec![unknown()])), pass_by: PassBy::Value,
            is_rest: false, is_kwargs: false, is_optional: true, is_nullable: false });
        assert!(inventory(object_member(kind, declaration)).by_binding.is_empty());
    }
}

#[test]
fn object_member_non_function_statement_values_reject_the_inventory() {
    for kind in 0..3 {
        assert!(inventory(object_member(kind, Statement::new(StmtKind::Return(None))))
            .by_binding.is_empty());
    }
}

#[test]
fn nested_object_method_expression_visits_do_not_rewalk_subtrees() {
    for depth in [16, 32, 64, 128] {
        let mut expression = Expression::int(0);
        for _ in 0..depth {
            expression = object_member(0, function("property_only", vec![root(expression)]));
        }
        let source = root(expression);
        let mut walk = BindingWalk::default();
        assert!(visit_binding_tree(&source, &mut walk, &mut |_| {}));
        assert_eq!(walk.expression_callbacks, depth + 1);
        assert_eq!(walk.expressions.len(), depth + 1);
        assert_eq!(walk.statements.len(), depth + 1);
    }
}

fn parameter_default(value: Expression) -> Param {
    Param { name: "bound".into(), type_hint: None, default: Some(value),
        pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
        is_optional: true, is_nullable: false }
}

fn pattern_default(kind: u8, value: Expression) -> BindingPattern {
    match kind {
        0 => BindingPattern::Object(vec![ObjectPatternProp { key: "bound".into(),
            value: None, default: Some(value), is_rest: false }]),
        1 => BindingPattern::Array(vec![ArrayPatternElem::Hole, ArrayPatternElem::Pattern(
            BindingPattern::Ident("bound".into()), Some(value)), ArrayPatternElem::Rest("rest".into())]),
        2 => BindingPattern::Object(vec![ObjectPatternProp { key: "outer".into(),
            value: Some(BindingPattern::Array(vec![ArrayPatternElem::Pattern(pattern_default(0, value), None)])),
            default: None, is_rest: false }]),
        _ => unreachable!(),
    }
}

fn audited_expression_slot(kind: u8, value: Expression) -> Expression {
    let zero = || Expression::int(0);
    Expression::new(match kind {
        0 => ExprKind::ArrayMap { array: Box::new(value), params: vec![], body: Box::new(zero()) },
        1 => ExprKind::ArrayMap { array: Box::new(zero()), params: vec![], body: Box::new(value) },
        2 => ExprKind::ArrayMap { array: Box::new(zero()), params: vec![parameter_default(value)], body: Box::new(zero()) },
        3 => ExprKind::ArrayTransform { op: ArrayTransformOp::MergeMask, args: vec![value], order: ArrayTraversalOrder::RowMajor },
        4 => ExprKind::WasmCallWithTag { tag: "tag".into(), callee: Box::new(value), args: vec![], tail: false, table: None },
        5 => ExprKind::WasmCallWithTag { tag: "tag".into(), callee: Box::new(zero()), args: vec![value], tail: true, table: Some(0) },
        6 => ExprKind::RefOf(Box::new(PlaceExpr::Deref(Box::new(value)))),
        7 => ExprKind::RefOf(Box::new(PlaceExpr::Member { object: Box::new(value), field: "field".into(), null_safe: false })),
        8 => ExprKind::RefOf(Box::new(PlaceExpr::Index { object: Box::new(value), index: Box::new(zero()), null_safe: false })),
        9 => ExprKind::RefOf(Box::new(PlaceExpr::Index { object: Box::new(zero()), index: Box::new(value), null_safe: false })),
        10..=12 => ExprKind::Destructure(match pattern_default(kind - 10, value) {
            BindingPattern::Object(fields) => DestructurePattern::Object(fields),
            BindingPattern::Array(elements) => DestructurePattern::Array(elements),
            _ => unreachable!(),
        }),
        13 => ExprKind::CallableRef { target: Box::new(Expression::ident("other")), receiver: None,
            binding: CallableBinding::Static, adapter: Some(CallableAdapter::Expr {
                params: vec![parameter_default(value)], body: Box::new(zero()),
            }) },
        14..=16 => ExprKind::Async(AsyncOp::Continue {
            source: Box::new(if kind == 14 { value.clone() } else { zero() }),
            on_fulfilled: Some(Box::new(if kind == 15 { value.clone() } else { zero() })),
            on_rejected: Some(Box::new(if kind == 16 { value } else { zero() })),
        }),
        17..=18 => ExprKind::Chan(ChanOp::New {
            capacity: Some(Box::new(if kind == 17 { value.clone() } else { zero() })),
            zero: Box::new(if kind == 18 { value } else { zero() }),
        }),
        19..=21 => ExprKind::Atomic(AtomicOp::CompareExchange {
            place: Box::new(if kind == 19 { value.clone() } else { zero() }),
            expected: Box::new(if kind == 20 { value.clone() } else { zero() }),
            replacement: Box::new(if kind == 21 { value } else { zero() }),
            result: RmwResult::Old, ordering: MemoryOrder::SeqCst,
        }),
        _ => unreachable!(),
    })
}

#[test]
fn audited_expression_slots_keep_unrelated_candidates_and_claim_nested_aliases() {
    for kind in 0..22 {
        assert_eq!(inventory(audited_expression_slot(kind, Expression::int(0))).by_binding.get("target"), Some(&0), "slot={kind}");
        let source = vec![function("Outer.Inner", vec![]),
            root(audited_expression_slot(kind, function_expr("Outer__Inner", vec![])))];
        let bindings = Bindings::prepare_with_aliases(&source, str::to_owned, aliases);
        assert_eq!(bindings.claimed.get("__php_fn_Outer__Inner"), Some(&None), "slot={kind}");
    }
}

#[test]
fn audited_expression_slots_reject_hidden_writes_and_unknown_statement_containers() {
    for kind in 0..22 {
        let write = Expression::new(ExprKind::Assign { target: Box::new(Expression::ident("target")), value: Box::new(Expression::int(0)) });
        assert!(!inventory(audited_expression_slot(kind, write)).by_binding.contains_key("target"), "write slot={kind}");
        let unknown = function_expr("guest", vec![Statement::new(StmtKind::NamespaceDecl { name: "opaque".into(), body: vec![] })]);
        assert!(inventory(audited_expression_slot(kind, unknown)).by_binding.is_empty(), "unknown slot={kind}");
    }
}

fn audited_statement_slot(kind: u8, value: Expression) -> Statement {
    Statement::new(match kind {
        0 => StmtKind::Break(BreakTarget::Value(value)),
        1..=4 => StmtKind::Switch { expr: Expression::int(0), cases: vec![SwitchCase {
            conditions: vec![match kind {
                1 => CaseCondition::Value(value),
                2 => CaseCondition::Range { from: value, to: Expression::int(10) },
                3 => CaseCondition::Range { from: Expression::int(0), to: value },
                4 => CaseCondition::Comparison { op: ComparisonOp::Eq, expr: value },
                _ => unreachable!(),
            }], body: vec![],
        }], default: None },
        5..=7 => StmtKind::VarDecl { declarations: vec![VarDeclarator {
            pattern: pattern_default(kind - 5, value), type_hint: None, init: None,
            array_bounds: None, with_events: false,
        }], kind: VarDeclKind::Let },
        8 => StmtKind::VarDecl { declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("bound".into()), type_hint: None, init: None,
            array_bounds: Some(vec![value]), with_events: false,
        }], kind: VarDeclKind::Let },
        9 => StmtKind::VarDecl { declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("bound".into()), type_hint: None, init: Some(value),
            array_bounds: None, with_events: false,
        }], kind: VarDeclKind::Let },
        10 => StmtKind::ClassDecl { name: "Container".into(), parents: vec![], interfaces: vec![], members: vec![],
            modifiers: Default::default(), decorators: vec![value] },
        11 => StmtKind::StructDecl { name: "Container".into(), interfaces: vec![], members: vec![],
            visibility: Visibility::Public, decorators: vec![value], semantics: Default::default() },
        12 => StmtKind::ClassDecl { name: "Container".into(), parents: vec![], interfaces: vec![],
            members: vec![ClassMember::Property { name: "property".into(), type_hint: None, getter: None,
                setter: Some(PropertySetter { param: parameter_default(value), body: vec![] }),
                is_auto: false, modifiers: Default::default() }], modifiers: Default::default(), decorators: vec![] },
        _ => unreachable!(),
    })
}

#[test]
fn audited_statement_slots_keep_unrelated_candidates_and_claim_nested_aliases() {
    for kind in 0..13 {
        let source = vec![function("target", vec![]), audited_statement_slot(kind, Expression::int(0))];
        assert_eq!(Bindings::prepare(&source, str::to_owned).by_binding.get("target"), Some(&0), "slot={kind}");
        let source = vec![function("Outer.Inner", vec![]), audited_statement_slot(kind, function_expr("Outer__Inner", vec![]))];
        let bindings = Bindings::prepare_with_aliases(&source, str::to_owned, aliases);
        assert_eq!(bindings.claimed.get("__php_fn_Outer__Inner"), Some(&None), "slot={kind}");
    }
}

#[test]
fn audited_statement_slots_reject_hidden_writes_and_unknown_statement_containers() {
    for kind in 0..13 {
        let write = Expression::new(ExprKind::Assign { target: Box::new(Expression::ident("target")), value: Box::new(Expression::int(0)) });
        let source = vec![function("target", vec![]), audited_statement_slot(kind, write)];
        assert!(!Bindings::prepare(&source, str::to_owned).by_binding.contains_key("target"), "write slot={kind}");
        let unknown = function_expr("guest", vec![Statement::new(StmtKind::NamespaceDecl { name: "opaque".into(), body: vec![] })]);
        let source = vec![function("target", vec![]), audited_statement_slot(kind, unknown)];
        assert!(Bindings::prepare(&source, str::to_owned).by_binding.is_empty(), "unknown slot={kind}");
    }
}

#[test]
fn walrus_writes_are_not_treated_as_stable_source_function_bindings() {
    let expression = Expression::new(ExprKind::Walrus { target: Box::new(Expression::ident("target")), value: Box::new(Expression::int(0)) });
    assert!(!inventory(expression).by_binding.contains_key("target"));
}

#[test]
fn omitted_expression_container_chains_have_linear_raw_callback_counts() {
    for depth in [16, 32, 64, 128] {
        for kind in [1, 3, 4] {
            let mut expression = Expression::int(0);
            for _ in 0..depth { expression = audited_expression_slot(kind, expression); }
            let source = root(expression);
            let mut walk = BindingWalk::default();
            assert!(visit_binding_tree(&source, &mut walk, &mut |_| {}));
            assert_eq!(walk.expression_callbacks, if kind == 1 { depth * 2 + 1 } else { depth + 1 });
            assert_eq!(walk.expressions.len(), walk.expression_callbacks);
        }
    }
}
