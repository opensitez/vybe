use super::*;
use vybe_ast::*;

fn parameter(name: &str) -> Param {
    Param { name: name.into(), type_hint: Some("number".into()), default: None,
        pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
        is_optional: false, is_nullable: false }
}
fn function(name: &str, params: Vec<Param>, return_type: Option<&str>, generator: bool) -> Statement {
    Statement::new(StmtKind::FunctionDecl { name: name.into(), params,
        return_type: return_type.map(str::to_owned), body: vec![Statement::new(StmtKind::Empty)],
        modifiers: Default::default(), handles: Vec::new(), is_async: false,
        is_generator: generator, is_sub: false })
}

fn oracle(statements: &[Statement]) -> Vec<FunctionDeclaration> {
    statements.iter().filter_map(|statement| {
        let StmtKind::FunctionDecl { name, params, return_type, is_generator, .. } = &statement.kind else { return None; };
        Some(FunctionDeclaration { name: name.clone(),
            param_modes: params.iter().map(|param| param.pass_by).collect(),
            param_types: params.iter().map(|param| param.type_hint.as_deref().map(str::to_string)).collect(),
            min_arity: params.iter().take_while(|param| param.default.is_none() && !param.is_rest).count(),
            signature: CallSignature::from_params(params), return_type: return_type.clone(),
            is_generator: *is_generator })
    }).collect()
}

fn assert_splits(body: &[Statement]) {
    let original = format!("{body:?}");
    let expected = format!("{:?}", oracle(body));
    for mask in 0..(1usize << body.len().saturating_sub(1)) {
        let mut preparation = FunctionRegistrationPreparation::default();
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
        assert_eq!(format!("{:?}", preparation.finish()), expected, "partition={mask}");
    }
    assert_eq!(format!("{body:?}"), original);
}

#[test]
fn staged_function_headers_preserve_defaults_modes_rest_kwargs_and_signature_names() {
    let mut default = parameter("$default");
    default.default = Some(Expression::with_span(ExprKind::Ident("later_global".into()),
        Span { start_line: 17, start_col: 2, end_line: 17, end_col: 14 }));
    default.is_optional = true;
    let mut reference = parameter("reference");
    reference.pass_by = PassBy::Ref;
    let mut rest = parameter("rest");
    rest.is_rest = true;
    let mut kwargs = parameter("kwargs");
    kwargs.is_kwargs = true;
    let body = vec![function("defaults", vec![parameter("required"), default, reference], Some("answer"), false),
        function("rest_and_kwargs", vec![parameter("required"), rest, kwargs.clone()], None, true),
        function("kwargs_only", vec![kwargs], None, false), function("empty", vec![], None, false)];
    let facts = oracle(&body);
    assert_eq!(facts[0].min_arity, 1);
    assert_eq!(facts[0].param_modes[2], PassBy::Ref);
    assert_eq!(facts[0].param_types[0].as_deref(), Some("number"));
    assert_eq!(facts[0].return_type.as_deref(), Some("answer"));
    assert_eq!(facts[0].signature.param_defaults[1].as_ref().unwrap().span.start_line, 17);
    assert_eq!(facts[0].signature.param_names[1], "default");
    assert_eq!(facts[1].signature.rest_index, Some(1));
    assert!(facts[1].signature.has_kwargs);
    assert!(facts[1].is_generator);
    assert!(!facts[1].signature.has_rest);
    assert_eq!((facts[2].min_arity, facts[2].signature.min_arity), (1, 0));
    assert_splits(&body);
}

#[test]
fn staged_function_headers_keep_duplicates_uncanonicalized_and_ordered() {
    let body = vec![function("Mixed", vec![parameter("first")], None, false),
        function("mixed", vec![parameter("second")], Some("later_type"), true),
        function("Mixed", vec![], Some("last_type"), false)];
    let mut preparation = FunctionRegistrationPreparation::default();
    preparation.consume(&body);
    let facts = preparation.finish();
    assert_eq!(facts.iter().map(|fact| fact.name.as_str()).collect::<Vec<_>>(), ["Mixed", "mixed", "Mixed"]);
    assert_splits(&body);
}

#[test]
fn staged_function_headers_keep_existing_top_level_only_registration_scope() {
    let nested = function("nested", vec![], None, false);
    let body = vec![Statement::new(StmtKind::Block(vec![nested.clone()])),
        Statement::new(StmtKind::NamespaceDecl { name: "scope".into(), body: vec![nested.clone()] }),
        Statement::new(StmtKind::ClassDecl { name: "container".into(), parents: Vec::new(), interfaces: Vec::new(),
            members: vec![ClassMember::Method(Box::new(nested.clone()))], modifiers: Default::default(), decorators: Vec::new() }),
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::FunctionExpr(Box::new(nested))))),
        function("top", vec![], None, false)];
    assert_eq!(oracle(&body).len(), 1);
    assert_splits(&body);
}

#[test]
fn staged_function_headers_own_default_expressions_and_abandoned_batches_are_isolated() {
    fn owned_and_send<T: Send + 'static>() {}
    owned_and_send::<FunctionRegistrationPreparation>();
    let mut param = parameter("defaulted");
    param.default = Some(Expression::int(23));
    let body = vec![function("first", vec![param], Some("number"), false), function("later", vec![], None, false)];
    let expected = format!("{:?}", oracle(&body));
    let mut preparation = FunctionRegistrationPreparation::default();
    for statement in body { preparation.consume(std::slice::from_ref(&statement)); }
    assert_eq!(format!("{:?}", preparation.finish()), expected);
    let mut abandoned = FunctionRegistrationPreparation::default();
    abandoned.consume(&[function("abandoned", vec![], None, false)]);
    drop(abandoned);
    assert!(FunctionRegistrationPreparation::default().finish().is_empty());
    assert_splits(&[]);
}
