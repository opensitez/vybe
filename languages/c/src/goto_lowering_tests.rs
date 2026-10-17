use super::*;

fn stmt(kind: StmtKind) -> Statement { Statement::new(kind) }
fn label(name: &str) -> Statement { stmt(StmtKind::Label(name.into())) }
fn jump(name: &str) -> Statement { stmt(StmtKind::GoTo(name.into())) }
fn safe(names: &[&str]) -> HashSet<String> { names.iter().map(|name| (*name).into()).collect() }

fn declaration(name: &str, type_name: &str, initializer: Expression) -> Statement {
    stmt(StmtKind::VarDecl { declarations: vec![VarDeclarator {
        pattern: BindingPattern::Ident(name.into()), type_hint: Some(type_name.into()),
        init: Some(initializer), array_bounds: None, with_events: false,
    }], kind: VarDeclKind::Let })
}

fn call(name: &str) -> Expression {
    expression(ExprKind::Call { callee: Box::new(ident(name)), args: Vec::new(), optional: false })
}

#[test]
fn unknown_duplicate_and_unsupported_destinations_reject_without_consuming_input() {
    for (body, expected) in [
        (vec![jump("missing")], "unresolved"),
        (vec![jump("same"), label("same"), label("same")], "duplicate"),
        (vec![jump("hidden"), stmt(StmtKind::Try {
            body: vec![label("hidden")], catches: Vec::new(), else_body: None, finally: None,
        })], "unsupported"),
        (vec![jump("hidden"), stmt(StmtKind::Labeled {
            label: "control".into(), body: Box::new(stmt(StmtKind::Block(vec![label("hidden")]))),
        })], "unsupported"),
    ] {
        let before = format!("{body:?}");
        let result = lower_function_gotos(&body, &safe(&[]), &safe(&[]));
        let Err(error) = result else { panic!("invalid preparation was accepted"); };
        assert!(error.reason.contains(expected), "{error:?}");
        assert_eq!(format!("{body:?}"), before, "failed preparation must preserve fallback AST");
    }
}

#[test]
fn nested_function_labels_are_not_parent_goto_destinations() {
    let nested = stmt(StmtKind::FunctionDecl {
        name: "nested".into(), params: Vec::new(), return_type: None,
        body: vec![label("inside"), jump("inside")], modifiers: Default::default(),
        handles: Vec::new(), is_async: false, is_generator: false, is_sub: false,
    });
    assert_eq!(format!("{:?}", lower_function_gotos(std::slice::from_ref(&nested), &safe(&[]), &safe(&[])).unwrap()),
        format!("{:?}", vec![nested.clone()]));
    let result = lower_function_gotos(&[jump("inside"), nested], &safe(&[]), &safe(&[]));
    assert!(result.is_err(), "a parent goto must never enter a nested callable");
}

#[test]
fn binding_reservation_requires_proof_and_keeps_initializer_in_one_original_state() {
    let declaration = stmt(StmtKind::VarDecl {
        declarations: vec![VarDeclarator { pattern: BindingPattern::Ident("value".into()),
            type_hint: Some("int".into()),
            init: Some(expression(ExprKind::Call { callee: Box::new(ident("effect_once")), args: Vec::new(), optional: false })),
            array_bounds: None, with_events: false }], kind: VarDeclKind::Let,
    });
    let body = vec![jump("after"), declaration.clone(), label("after"), stmt(StmtKind::Return(Some(integer(0))))];
    assert!(lower_function_gotos(&body, &safe(&[]), &safe(&[])).is_err());
    let lowered = lower_function_gotos(&body, &safe(&["value"]), &safe(&[])).unwrap();
    let StmtKind::VarDecl { declarations, .. } = &lowered[0].kind else { panic!("binding reservation"); };
    assert!(declarations[0].init.is_none(), "reservation cannot execute initializer");
    let mut calls = 0;
    for statement in &lowered {
        statement.walk_exprs(&mut |value| {
            if matches!(&value.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "effect_once")) { calls += 1; }
        });
    }
    assert_eq!(calls, 1, "initializer is retained once at its graph node, never cloned into prelude");
    let duplicated = vec![jump("after"), declaration.clone(), declaration, label("after")];
    assert!(lower_function_gotos(&duplicated, &safe(&["value"]), &safe(&[])).is_err(),
        "unrenamed shadow bindings cannot share a dispatcher-local slot");
}

#[test]
fn dynamic_arrays_nonordinary_storage_and_const_lifetimes_decline_even_with_a_name_receipt() {
    for mode in 0..5 {
        let mut declaration = VarDeclarator { pattern: BindingPattern::Ident("value".into()),
            type_hint: Some("int".into()), init: Some(integer(1)), array_bounds: None, with_events: false };
        let kind = match mode {
            0 => { declaration.array_bounds = Some(vec![ident("dynamic_bound")]); VarDeclKind::Let }
            1 => { declaration.type_hint = Some("int[4]".into()); VarDeclKind::Let }
            2 => { declaration.type_hint.as_mut().unwrap().access = Some(StorageAccess::Volatile); VarDeclKind::Let }
            3 => VarDeclKind::Static,
            _ => VarDeclKind::Const,
        };
        let body = vec![jump("after"), stmt(StmtKind::VarDecl { declarations: vec![declaration], kind }), label("after")];
        assert!(lower_function_gotos(&body, &safe(&["value"]), &safe(&[])).is_err(), "mode={mode}");
    }
}

#[test]
fn expression_owned_unlowered_control_is_an_explicit_error_not_an_opaque_leaf() {
    let value = expression(ExprKind::Lambda { params: Vec::new(),
        body: LambdaBody::Block(vec![jump("inside"), label("inside")]),
        is_async: false, captures: Vec::new() });
    let body = vec![stmt(StmtKind::Expr(value))];
    assert!(lower_function_gotos(&body, &safe(&[]), &safe(&[])).is_err());
}

#[test]
fn no_goto_is_identity_and_generated_dispatch_names_respect_owner_reservations() {
    let body = vec![stmt(StmtKind::Return(Some(integer(3))))];
    assert_eq!(format!("{:?}", lower_function_gotos(&body, &safe(&[]), &safe(&[])).unwrap()), format!("{body:?}"));
    let body = vec![jump("target"), label("target"), stmt(StmtKind::Return(Some(integer(3))))];
    let lowered = lower_function_gotos(&body, &safe(&[]),
        &safe(&["__c_function_goto_pc_0", "__c_function_goto_dispatch_0"])).unwrap();
    let StmtKind::VarDecl { declarations, .. } = &lowered[0].kind else { panic!("pc declaration"); };
    assert!(matches!(&declarations[0].pattern, BindingPattern::Ident(name) if name == "__c_function_goto_pc_1"));
    assert!(!format!("{lowered:?}").contains("GoTo("));
}

#[test]
fn leading_initializer_stays_once_before_dispatch_and_outside_every_backward_edge() {
    let prefix = declaration("value", "int", call("prefix_once"));
    let body = vec![prefix.clone(), label("again"), stmt(StmtKind::If {
        cond: ident("repeat"), then_body: vec![jump("again")], elifs: Vec::new(), else_body: None,
    }), stmt(StmtKind::Return(Some(ident("value"))))];
    let lowered = lower_function_gotos(&body, &safe(&[]), &safe(&["repeat"])).unwrap();
    assert_eq!(format!("{:?}", lowered[0]), format!("{prefix:?}"));
    let mut all_calls = 0;
    let mut dispatch_calls = 0;
    for (index, statement) in lowered.iter().enumerate() {
        statement.walk_exprs(&mut |value| {
            if matches!(&value.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "prefix_once"))
            {
                all_calls += 1;
                if index > 0 { dispatch_calls += 1; }
            }
        });
    }
    assert_eq!((all_calls, dispatch_calls), (1, 0), "back edges can enter only dispatcher states, not the prefix");
    assert!(!format!("{lowered:?}").contains("GoTo("));
}

#[test]
fn pointer_aggregate_and_volatile_leading_declarations_remain_exact_without_split_receipts() {
    let mut prefix = vec![
        declaration("pointer", "int *", call("pointer_once")),
        declaration("record", "Record", call("record_once")),
        declaration("word", "int", call("volatile_once")),
    ];
    let StmtKind::VarDecl { declarations, .. } = &mut prefix[2].kind else { unreachable!(); };
    declarations[0].type_hint.as_mut().unwrap().access = Some(StorageAccess::Volatile);
    for (index, statement) in prefix.iter_mut().enumerate() {
        statement.span.start_line = index as u32 + 10;
        statement.span.end_line = index as u32 + 10;
    }
    let mut body = prefix.clone();
    body.extend([jump("inside"), stmt(StmtKind::If { cond: integer(0),
        then_body: vec![label("inside"), stmt(StmtKind::Return(Some(integer(0))))],
        elifs: Vec::new(), else_body: None })]);
    let lowered = lower_function_gotos(&body, &safe(&[]), &safe(&[])).unwrap();
    assert_eq!(format!("{:?}", &lowered[..prefix.len()]), format!("{prefix:?}"),
        "preserve storage, access, initializer, source order and spans; no scalar recategorization");
}

#[test]
fn prefix_names_are_reserved_but_late_declarations_still_need_unique_strong_proof() {
    let prefix = declaration("__c_function_goto_pc_0", "int", integer(7));
    let late = declaration("late", "int", call("late_once"));
    let body = vec![prefix.clone(), label("again"), late, jump("again")];
    assert!(lower_function_gotos(&body, &safe(&[]), &safe(&[])).is_err());
    let lowered = lower_function_gotos(&body, &safe(&["late"]), &safe(&[])).unwrap();
    assert_eq!(format!("{:?}", lowered[0]), format!("{prefix:?}"));
    let StmtKind::VarDecl { declarations, .. } = &lowered[2].kind else { panic!("pc declaration"); };
    assert!(matches!(&declarations[0].pattern, BindingPattern::Ident(name) if name == "__c_function_goto_pc_1"));

    let duplicate_late = vec![prefix.clone(), label("again"), prefix.clone(), jump("again")];
    assert!(lower_function_gotos(&duplicate_late, &safe(&["__c_function_goto_pc_0"]), &safe(&[])).is_err());
    let duplicate_prefix = vec![prefix.clone(), prefix, jump("after"), label("after")];
    assert!(lower_function_gotos(&duplicate_prefix, &safe(&[]), &safe(&[])).is_err());
}

#[test]
fn prefix_initializer_with_embedded_control_label_is_not_exempt_from_callable_guard() {
    let value = expression(ExprKind::Lambda { params: Vec::new(),
        body: LambdaBody::Block(vec![label("private"), stmt(StmtKind::Return(Some(integer(7))))]),
        is_async: false, captures: Vec::new() });
    let body = vec![declaration("value", "int", value), jump("after"), label("after")];
    let original = format!("{body:?}");
    assert!(lower_function_gotos(&body, &safe(&[]), &safe(&[])).is_err());
    assert_eq!(format!("{body:?}"), original);
}

