use super::*;

fn expression(kind: ExprKind) -> Expression { Expression::new(kind) }
fn ident(name: &str) -> Expression { expression(ExprKind::Ident(name.into())) }
fn stmt(kind: StmtKind) -> Statement { Statement::new(kind) }
fn address(name: &str) -> Statement {
    stmt(StmtKind::Expr(expression(ExprKind::Unary {
        op: UnaryOp::AddrOf, expr: Box::new(ident(name)),
    })))
}
fn declare(name: &str, kind: VarDeclKind) -> Statement {
    pattern(BindingPattern::Ident(name.into()), kind)
}
fn pattern(pattern: BindingPattern, kind: VarDeclKind) -> Statement {
    stmt(StmtKind::VarDecl { declarations: vec![VarDeclarator {
        pattern, type_hint: None, init: None, array_bounds: None, with_events: false,
    }], kind })
}
fn parameter(name: &str, default: Option<Expression>) -> Param {
    Param { name: name.into(), type_hint: None, default, pass_by: PassBy::Value,
        is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false }
}
fn function(name: &str, params: Vec<Param>, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::FunctionDecl { name: name.into(), params, return_type: None,
        body, modifiers: Default::default(), handles: Vec::new(),
        is_async: false, is_generator: false, is_sub: false })
}
fn scope(kind: ScopeDeclKind, names: &[&str]) -> Statement {
    stmt(StmtKind::ScopeDecl { kind, names: names.iter().map(|name| (*name).into()).collect() })
}

// Original entry point: prebind and traverse the entire module in one call.
// Keep this oracle independent of the staged entry's header/consume protocol.
fn whole_module_oracle(body: &[Statement], fold: Option<CaseAlphabet>) -> HashSet<String> {
    let mut scan = Scan { frames: vec![Frame { function: true, ..Frame::default() }],
        globals: HashSet::new(), fold };
    scan.body(body);
    scan.globals
}

fn partitions(body: &[Statement], mask: usize, mut visit: impl FnMut(&[Statement])) {
    let mut start = 0;
    visit(&[]);
    for end in 1..body.len() {
        if mask & (1 << (end - 1)) != 0 {
            visit(&body[start..end]);
            start = end;
        }
    }
    visit(&body[start..]);
    visit(&[]);
}

fn assert_splits(body: &[Statement], fold: Option<CaseAlphabet>) {
    let expected = whole_module_oracle(body, fold);
    let original = format!("{body:?}");
    let partition_count = 1usize << body.len().saturating_sub(1);
    for headers in 0..partition_count {
        for units in 0..partition_count {
            let mut bindings = ModuleReferenceBindings::new(fold);
            partitions(body, headers, |chunk| bindings.collect(chunk));
            let mut preparation = bindings.seal();
            partitions(body, units, |chunk| preparation.consume(chunk));
            assert_eq!(preparation.finish(), expected, "headers={headers} units={units} fold={fold:?}");
        }
    }
    assert_eq!(module_address_taken(body, fold), expected);
    assert_eq!(format!("{body:?}"), original);
}

#[test]
fn staged_reference_preparation_preserves_lexical_and_function_scope_boundaries() {
    let body = vec![
        function("before", vec![parameter("parameter", None)], vec![
            address("parameter"),
            stmt(StmtKind::Block(vec![address("lexical"), declare("lexical", VarDeclKind::Let)])),
            address("ordinary"), declare("ordinary", VarDeclKind::Var), address("ordinary"),
            address("hoisted"),
            stmt(StmtKind::Block(vec![declare("hoisted", VarDeclKind::FunctionScoped)])),
            address("future_global"),
        ]),
        declare("future_global", VarDeclKind::Const),
        stmt(StmtKind::Block(vec![address("block_local"), declare("block_local", VarDeclKind::Let)])),
        address("module_direct"),
    ];
    assert_eq!(whole_module_oracle(&body, None),
        ["ordinary", "future_global", "module_direct"].map(str::to_owned).into());
    for fold in [None, Some(CaseAlphabet::Ascii), Some(CaseAlphabet::Unicode)] { assert_splits(&body, fold); }
}

#[test]
fn staged_reference_preparation_preserves_closures_defaults_and_scope_redirects() {
    let default = expression(ExprKind::Unary { op: UnaryOp::AddrOf, expr: Box::new(ident("default_global")) });
    let body = vec![
        function("outer", vec![parameter("captured", None)], vec![
            function("inner", vec![parameter("defaulted", Some(default))], vec![
                scope(ScopeDeclKind::Nonlocal, &["captured"]), address("captured"),
                scope(ScopeDeclKind::Global, &["redirected"]), address("redirected"),
            ]),
            function("closed", vec![], vec![scope(ScopeDeclKind::Closed, &[]), address("hidden"),
                scope(ScopeDeclKind::Global, &["opened"]), address("opened")]),
        ]),
        stmt(StmtKind::Expr(expression(ExprKind::Lambda {
            params: vec![parameter("lambda_local", None)],
            body: LambdaBody::Block(vec![address("lambda_local"), address("lambda_global")]),
            is_async: false, captures: Vec::new(),
        }))),
        declare("redirected", VarDeclKind::Var),
        stmt(StmtKind::Empty),
    ];
    assert_eq!(whole_module_oracle(&body, None),
        ["default_global", "redirected", "opened", "lambda_global"].map(str::to_owned).into());
    for fold in [None, Some(CaseAlphabet::Ascii), Some(CaseAlphabet::Unicode)] { assert_splits(&body, fold); }
}

#[test]
fn staged_reference_preparation_preserves_namespaces_patterns_and_reference_forms() {
    let destructuring = pattern(BindingPattern::Object(vec![ObjectPatternProp {
        key: "local".into(), value: Some(BindingPattern::Array(vec![
            ArrayPatternElem::Pattern(BindingPattern::Ident("element".into()), None),
            ArrayPatternElem::Rest("rest".into()),
        ])), default: None, is_rest: false,
    }]), VarDeclKind::Let);
    let mut by_ref = Argument::positional(ident("argument_global"));
    by_ref.by_ref = true;
    let body = vec![
        stmt(StmtKind::NamespaceDecl { name: "scope".into(), body: vec![
            function("nested", vec![], vec![address("element"), address("rest"), destructuring]),
            address("namespace_global"),
        ] }),
        stmt(StmtKind::Expr(expression(ExprKind::RefOf(Box::new(PlaceExpr::Ident("place_global".into())))))),
        stmt(StmtKind::Expr(expression(ExprKind::Call {
            callee: Box::new(ident("callee")), args: vec![by_ref], optional: false,
        }))),
        stmt(StmtKind::Assign { targets: vec![ident("alias")], value: ident("assigned_global"), by_ref: true }),
    ];
    assert_eq!(whole_module_oracle(&body, None),
        ["namespace_global", "place_global", "argument_global", "assigned_global"].map(str::to_owned).into());
    for fold in [None, Some(CaseAlphabet::Ascii), Some(CaseAlphabet::Unicode)] { assert_splits(&body, fold); }
}

#[test]
fn staged_reference_preparation_preserves_ascii_and_unicode_folding() {
    let body = vec![function("folding", vec![parameter("LOCAL", None), parameter("\u{e4}name", None)],
        vec![address("local"), address("\u{c4}NAME"), address("FrEe")]),
        declare("later", VarDeclKind::Let), address("later")];
    for (fold, names) in [
        (None, vec!["local", "\u{c4}NAME", "FrEe", "later"]),
        (Some(CaseAlphabet::Ascii), vec!["\u{c4}name", "free", "later"]),
        (Some(CaseAlphabet::Unicode), vec!["free", "later"]),
    ] {
        assert_eq!(whole_module_oracle(&body, fold), names.into_iter().map(str::to_owned).collect());
        assert_splits(&body, fold);
    }
}

#[test]
fn staged_reference_facts_outlive_consumed_statements_and_abandoned_sources() {
    fn owned_and_send<T: Send + 'static>() {}
    owned_and_send::<ModuleReferenceBindings>();
    owned_and_send::<ModuleReferencePreparation>();
    let body = vec![address("owned_fact"), declare("global", VarDeclKind::Let),
        function("f", vec![], vec![address("other_fact")])];
    let expected = whole_module_oracle(&body, None);
    let mut bindings = ModuleReferenceBindings::new(None);
    for statement in &body { bindings.collect(std::slice::from_ref(statement)); }
    let mut preparation = bindings.seal();
    for statement in body {
        preparation.consume(std::slice::from_ref(&statement));
        // Each complete source unit is dropped here, before facts are sealed.
    }
    assert_eq!(preparation.finish(), expected);
    let mut abandoned = ModuleReferenceBindings::new(None).seal();
    abandoned.consume(&[address("failed_source")]);
    drop(abandoned);
    assert!(ModuleReferenceBindings::new(None).seal().finish().is_empty());
    assert_splits(&[], None);
}
