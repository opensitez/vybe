use super::*;
use vybe_ast::*;

fn value(kind: ExprKind) -> Expression { Expression::new(kind) }
fn statement(kind: StmtKind) -> Statement { Statement::new(kind) }
fn bigint() -> Statement { statement(StmtKind::Expr(value(ExprKind::Lit(Literal::BigInt(7))))) }
fn proxy() -> Expression {
    value(ExprKind::Proxy { target: Box::new(Expression::ident("target")),
        handler: Box::new(Expression::ident("handler")) })
}
fn function(body: Vec<Statement>) -> Statement {
    statement(StmtKind::FunctionDecl { name: "procedure".into(), params: Vec::new(), return_type: None,
        body, modifiers: Default::default(), handles: Vec::new(),
        is_async: false, is_generator: false, is_sub: false })
}

fn oracle(body: &[Statement], bigint_enabled: bool, inspect_proxy: bool, uses_proxy: bool) -> SourceFeatures {
    let bigint_enabled = bigint_enabled || body.iter().any(|statement| {
        let mut found = false;
        statement.walk_exprs(&mut |value| { found |= matches!(value.kind, ExprKind::Lit(Literal::BigInt(_))); });
        found
    });
    let uses_proxy = uses_proxy || (inspect_proxy && body.iter().any(stmt_uses_proxy));
    SourceFeatures { bigint_enabled, uses_proxy }
}

fn assert_splits(body: &[Statement]) {
    let original = format!("{body:?}");
    for bigint_enabled in [false, true] {
        for inspect_proxy in [false, true] {
            for uses_proxy in [false, true] {
                let expected = oracle(body, bigint_enabled, inspect_proxy, uses_proxy);
                for mask in 0..(1usize << body.len().saturating_sub(1)) {
                    let mut preparation = SourceFeaturePreparation::new(bigint_enabled, inspect_proxy, uses_proxy);
                    preparation.consume(&[]);
                    let mut start = 0;
                    for end in 1..body.len() {
                        if mask & (1 << (end - 1)) != 0 {
                            preparation.consume(&body[start..end]);
                            start = end;
                        }
                    }
                    preparation.consume(&body[start..]);
                    preparation.consume(&[]);
                    assert_eq!(preparation.finish(), expected, "partition={mask}");
                }
            }
        }
    }
    assert_eq!(format!("{body:?}"), original);
}

#[test]
fn staged_source_features_preserve_policy_seeds_and_late_discovery() {
    let body = vec![statement(StmtKind::Empty), function(vec![statement(StmtKind::Empty)]),
        bigint(), statement(StmtKind::Expr(proxy())), statement(StmtKind::Empty)];
    assert_eq!(oracle(&body, false, true, false), SourceFeatures { bigint_enabled: true, uses_proxy: true });
    assert_eq!(oracle(&body, false, false, false), SourceFeatures { bigint_enabled: true, uses_proxy: false });
    assert_splits(&body);
    assert_splits(&[]);
}

#[test]
fn staged_source_features_preserve_nested_callable_and_revocable_coverage() {
    let revocable = value(ExprKind::Call { callee: Box::new(value(ExprKind::Member {
        object: Box::new(Expression::ident("Proxy")), field: "revocable".into(), null_safe: false,
    })), args: Vec::new(), optional: false });
    for creation in [proxy(), revocable] {
        let body = vec![function(vec![statement(StmtKind::Return(Some(value(ExprKind::Lambda {
            params: Vec::new(), body: LambdaBody::Expr(Box::new(creation)),
            is_async: false, captures: Vec::new(),
        }))))]), bigint()];
        assert!(oracle(&body, false, true, false).uses_proxy);
        assert_splits(&body);
    }
}

#[test]
fn staged_source_features_do_not_broaden_existing_visitor_or_spelling_policy() {
    let body = vec![statement(StmtKind::Expr(value(ExprKind::New {
        class: Box::new(Expression::ident("Proxy")), args: Vec::new(),
    }))), statement(StmtKind::Expr(value(ExprKind::Lit(Literal::Str("BigInt Proxy".into())))))];
    assert_eq!(oracle(&body, false, true, false), SourceFeatures { bigint_enabled: false, uses_proxy: false });
    assert_splits(&body);
    // The two original visitors intentionally have different syntax coverage.
    // Staging must preserve it, not replace both with a new broad traversal.
    let namespace = vec![statement(StmtKind::NamespaceDecl { name: "scope".into(),
        body: vec![bigint(), statement(StmtKind::Expr(proxy()))] })];
    assert_eq!(oracle(&namespace, false, true, false), SourceFeatures { bigint_enabled: true, uses_proxy: false });
    assert_splits(&namespace);
}

#[test]
fn staged_source_feature_facts_outlive_units_and_abandoned_sources_are_isolated() {
    fn owned_and_send<T: Send + 'static>() {}
    owned_and_send::<SourceFeaturePreparation>();
    owned_and_send::<SourceFeatures>();
    let mut preparation = SourceFeaturePreparation::new(false, true, false);
    for unit in [bigint(), statement(StmtKind::Expr(proxy()))] {
        preparation.consume(std::slice::from_ref(&unit));
    }
    assert_eq!(preparation.finish(), SourceFeatures { bigint_enabled: true, uses_proxy: true });
    let mut abandoned = SourceFeaturePreparation::new(false, true, false);
    abandoned.consume(&[bigint(), statement(StmtKind::Expr(proxy()))]);
    drop(abandoned);
    assert_eq!(SourceFeaturePreparation::new(false, true, false).finish(),
        SourceFeatures { bigint_enabled: false, uses_proxy: false });
}
