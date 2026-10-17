use super::*;

// Previous implementation retained as an independent equivalence oracle.
fn restoring_binder(mut expression: Expression, walker: &mut PhpWalker) -> Expression {
    if !walker.function_symbol_scope_stack.last().is_some_and(Option::is_some) { return expression; }
    let mut boundaries = std::collections::HashMap::new();
    expression.walk_exprs(&mut |expr| {
        if matches!(expr.kind, ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_)) {
            boundaries.insert(expr as *const Expression as usize, expr.kind.clone());
        }
    });
    expression.walk_exprs_mut(&mut |expr| {
        if let Some(original) = boundaries.remove(&(expr as *const Expression as usize)) { expr.kind = original; return; }
        let ExprKind::Ident(name) = &expr.kind else { return; };
        if !name.starts_with('$') || name.starts_with("$__php_") || name == "$this" || name == "$GLOBALS"
            || PHP_SUPERGLOBAL_NAMES.contains(&name.as_str()) || php_is_function_global(walker, name) { return; }
        if let Some(bound) = php_symbol_scope_variable(walker, Expression::string(name.trim_start_matches('$')), &expr.span) { *expr = bound; }
    });
    expression
}

fn scope_walker(lhs:u32,enabled:bool)->PhpWalker {
    let mut walker=PhpWalker::default();
    walker.assign_lhs_depth=lhs;
    walker.function_symbol_scope_stack.push(enabled.then(||"$__php_scope_test".into()));
    walker.function_global_stack.push(["global".to_owned()].into_iter().collect());
    walker
}

fn scope_fixture(depth:usize)->Expression {
    let mut values=["$local","$global","$this","$GLOBALS","$_GET","$_POST","$__php_temp","bare"]
        .map(Expression::ident).to_vec();
    if depth>0 {
        values.push(Expression::new(ExprKind::Lambda {
            params:vec![],body:LambdaBody::Expr(Box::new(scope_fixture(depth-1))),
            is_async:false,captures:vec!["$local".into()],
        }));
        values.push(Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(StmtKind::FunctionDecl {
            name:"nested".into(),params:vec![],return_type:None,
            body:vec![Statement::new(StmtKind::Expr(scope_fixture(depth-1)))],
            modifiers:Modifiers::default(),handles:vec![],is_async:false,is_generator:true,is_sub:false,
        })))));
    }
    Expression::with_span(ExprKind::Tuple(values),Span{start_line:7,start_col:2,end_line:9,end_col:4})
}

#[test]
fn pruned_symbol_binding_preserves_ast_in_read_write_and_unbound_scopes(){
    for enabled in [false,true] {for lhs in [0,1,2] {for depth in 0..5 {
        let expression=scope_fixture(depth);
        let expected=restoring_binder(expression.clone(),&mut scope_walker(lhs,enabled));
        let actual=php_bind_symbol_scope_expression(expression,&mut scope_walker(lhs,enabled));
        assert_eq!(format!("{actual:?}"),format!("{expected:?}"),"enabled={enabled} lhs={lhs} depth={depth}");
    }}}
}

#[test]
fn nested_scope_body_keeps_its_owned_storage(){
    let expression=scope_fixture(3);
    let ExprKind::Tuple(values)=&expression.kind else{unreachable!()};
    let ExprKind::Lambda{body:LambdaBody::Expr(body),..}=&values[8].kind else{unreachable!()};
    let original_body=body.as_ref() as *const Expression;
    let ExprKind::FunctionExpr(declaration)=&values[9].kind else{unreachable!()};
    let original_declaration=declaration.as_ref() as *const Statement;
    let actual=php_bind_symbol_scope_expression(expression,&mut scope_walker(0,true));
    let ExprKind::Tuple(values)=actual.kind else{unreachable!()};
    assert!(matches!(values[0].kind,ExprKind::RefLoad(_)));
    assert!(matches!(&values[1].kind,ExprKind::Ident(name) if name=="$global"));
    let ExprKind::Lambda{body:LambdaBody::Expr(body),..}=&values[8].kind else{unreachable!()};
    assert_eq!(body.as_ref() as *const Expression,original_body);
    let ExprKind::FunctionExpr(declaration)=&values[9].kind else{unreachable!()};
    assert_eq!(declaration.as_ref() as *const Statement,original_declaration);
}

#[test]
fn live_scope_reference_argument_uses_cell_lookup_not_value_index() {
    let mut walker = scope_walker(0, true);
    let value = php_symbol_scope_variable(&walker, Expression::string("text"), &Span::default()).unwrap();
    let reference = php_reference_argument_value(&mut walker, &value);
    let ExprKind::Unary { op: UnaryOp::Deref, expr } = reference.kind else { panic!("expected reference place") };
    let ExprKind::Call { callee, args, .. } = expr.kind else { panic!("expected cell lookup") };
    assert!(matches!(callee.kind, ExprKind::Ident(ref name) if name == "__php_scope_ref"));
    assert_eq!(args.len(), 2);
    assert!(matches!(&args[0].value.kind, ExprKind::Ident(name) if name == "$__php_scope_test"));
    assert!(matches!(&args[1].value.kind, ExprKind::Lit(Literal::Str(key)) if key == "text"));
}

#[test]
fn reference_pointer_preserves_existing_pointer_expressions() {
    let walker = scope_walker(0, true);
    for pointer in [Expression::ident("$pointer"), php_mk_call("__php_scope_ref",
        vec![Expression::ident("$__php_scope_test"), Expression::string("text")], &Span::default())] {
        assert_eq!(format!("{:?}", php_reference_pointer(&walker, &pointer)), format!("{pointer:?}"));
    }
}

#[test]
fn array_literal_reference_tokens_become_places() {
    for (rule, source) in [
        (Rule::short_array_expression, "[&$value]"),
        (Rule::array_expression, "array('key' => /* gap */ &$value)"),
        (Rule::short_array_expression, "['key' => & /* gap */ $value]"),
    ] {
        let mut walker = PhpWalker::default();
        let pair = PhpParser::parse(rule, source).unwrap().next().unwrap();
        let array = walk_array(&mut walker, pair).unwrap();
        let ExprKind::Array(elements) = array.kind else { panic!("expected array") };
        assert_eq!(elements.len(), 1, "{source}");
        assert!(elements[0].by_ref, "{source}");
        assert!(matches!(elements[0].value.kind, ExprKind::RefOf(_) |
            ExprKind::Unary { op: UnaryOp::AddrOf, .. }), "{source}");
    }
}

#[test]
fn array_reference_gap_ignores_ampersands_in_comments() {
    for gap in [" => /* & */ ", " // &\n ", " # &\n "] {
        assert!(!php_array_gap_has_reference(gap), "{gap:?}");
    }
    for gap in [" & ", " => /* & */ & ", " & /* gap */ "] {
        assert!(php_array_gap_has_reference(gap), "{gap:?}");
    }
}
