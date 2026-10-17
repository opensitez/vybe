use super::*;

fn roots_for(expression: Expression) -> PhpBootstrapNeeds {
    php_bootstrap_needs(&[Statement::new(StmtKind::Expr(expression))])
}

#[test]
fn bootstrap_resolved_implode_contract_keeps_typeerror_ancestry() {
    for name in ["implode", "join", "IMPlode", "\\implode"] {
        let needs = roots_for(php_call_expr(name, vec![Expression::string("-"), Expression::new(ExprKind::Array(vec![]))]));
        assert!(!needs.exceptions, "{name}");
        assert_eq!(needs.generated_exception_roots, ["TypeError"], "{name}");
        let names: Vec<_> = crate::core_exceptions::declarations_for(&needs.generated_exception_roots)
            .into_iter().filter_map(|stmt| if let StmtKind::ClassDecl { name, .. } = stmt.kind { Some(name) } else { None }).collect();
        assert_eq!(names, ["Throwable", "Error", "TypeError"]);
    }
}

#[test]
fn bootstrap_implode_contract_declines_unresolved_shapes() {
    let value = Argument::positional(Expression::ident("$array"));
    for name in ["Project.implode", "Project\\implode", "$callable", "other"] {
        assert!(php_resolved_builtin_exception_roots(&Expression::ident(name), &[value.clone()]).is_none());
    }
    assert!(php_resolved_builtin_exception_roots(&Expression::ident("implode"), &[]).is_none());
    assert!(php_resolved_builtin_exception_roots(&Expression::ident("implode"), &[value.clone(), value.clone(), value.clone()]).is_none());
    for argument in [Argument { spread: true, ..value.clone() }, Argument { by_ref: true, ..value.clone() },
        Argument { name: Some("array".into()), ..value.clone() }] {
        assert!(php_resolved_builtin_exception_roots(&Expression::ident("implode"), &[argument]).is_none());
    }
    let member = Expression::new(ExprKind::Member { object: Box::new(Expression::ident("$object")), field: "implode".into(), null_safe: false });
    assert!(php_resolved_builtin_exception_roots(&member, &[value]).is_none());
}

#[test]
fn bootstrap_implode_contract_does_not_hide_nested_or_dynamic_needs() {
    let needs = roots_for(php_call_expr("implode", vec![php_call_expr("unknown", vec![])]));
    assert!(needs.generated_exception_roots.contains(&"PDOException"));
    for name in ["class_exists", "get_declared_classes", "__php_dynamic_new"] {
        let needs = roots_for(php_call_expr(name, vec![Expression::string("PDOException")]));
        assert!(needs.generated_exception_roots.contains(&"PDOException"), "{name}");
    }
    let needs = roots_for(Expression::new(ExprKind::New { class: Box::new(Expression::ident("$class")), args: vec![] }));
    assert!(needs.generated_exception_roots.contains(&"PDOException"));
    let needs = roots_for(Expression::new(ExprKind::New { class: Box::new(Expression::ident("Exception")), args: vec![] }));
    assert!(needs.exceptions);
}

#[test]
fn bootstrap_exact_array_contracts_keep_typeerrors_and_exclude_search_arities() {
    for name in ["array_keys", "ARRAY_VALUES"] {
        let needs = roots_for(php_call_expr(name, vec![Expression::ident("$input")]));
        assert_eq!(needs.generated_exception_roots, ["TypeError"], "{name}");
    }
    let arg = Argument::positional(Expression::ident("$input"));
    for args in [vec![arg.clone(), arg.clone()], vec![arg.clone(), arg.clone(), arg.clone()]] {
        assert!(php_resolved_builtin_exception_roots(&Expression::ident("array_keys"), &args).is_none());
    }
    for name in ["array_is_list", "in_array", "count", "array_map", "Project.array_values"] {
        assert!(php_resolved_builtin_exception_roots(&Expression::ident(name), &[arg.clone()]).is_none(), "{name}");
    }
}

#[test]
fn bootstrap_total_is_array_has_no_generated_roots_but_scans_argument_calls() {
    for value in [Expression::ident("$object"), Expression::int(1), Expression::bool(false), Expression::null()] {
        let needs = roots_for(php_call_expr("is_array", vec![value]));
        assert!(!needs.exceptions);
        assert!(needs.generated_exception_roots.is_empty());
    }
    let needs = roots_for(php_call_expr("is_array", vec![php_call_expr("unknown", vec![])]));
    assert!(needs.generated_exception_roots.contains(&"PDOException"));
    assert!(php_resolved_builtin_exception_roots(&Expression::ident("is_array"), &[]).is_none());
}
