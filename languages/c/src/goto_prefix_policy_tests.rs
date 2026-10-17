use super::*;
use vybe_ast::{BindingPattern, ExprKind, Literal, VarDeclKind, VarDeclarator};

fn stmt(kind: StmtKind) -> Statement { Statement::new(kind) }
fn number(value: i64) -> vybe_ast::Expression { vybe_ast::Expression::new(ExprKind::Lit(Literal::Int(value))) }
fn go(name: &str) -> Statement { stmt(StmtKind::GoTo(name.into())) }
fn label(name: &str) -> Statement { stmt(StmtKind::Label(name.into())) }
fn declaration(name: &str, hint: &str) -> Statement {
    stmt(StmtKind::VarDecl { declarations: vec![VarDeclarator {
        pattern: BindingPattern::Ident(name.into()), type_hint: Some(hint.into()),
        init: Some(number(0)), array_bounds: None, with_events: false,
    }], kind: VarDeclKind::Let })
}

#[test]
fn prefix_graph_admits_unmoved_nonscalar_prefix_without_any_binding_grants() {
    let prefix = declaration("pointer", "int *");
    let body = vec![prefix.clone(), go("nested"), stmt(StmtKind::If { cond: number(0),
        then_body: vec![label("nested"), stmt(StmtKind::Return(Some(number(0))))],
        elifs: Vec::new(), else_body: None })];
    let result = prepare(&body, &HashSet::new()).unwrap();
    assert_eq!(result.path, Path::PrefixGraph);
    assert_eq!(format!("{:?}", result.body[0]), format!("{prefix:?}"));
    assert!(!contains_unprepared_control(&result.body, false));
}

#[test]
fn late_declaration_uses_legacy_only_when_every_destination_is_top_level() {
    let body = vec![go("after"), declaration("late", "int"), label("after"),
        stmt(StmtKind::Return(Some(number(0))))];
    let result = prepare(&body, &HashSet::new()).unwrap();
    assert_eq!(result.path, Path::ProvenTopTargets);
    assert!(!contains_unprepared_control(&result.body, false));
}

#[test]
fn known_top_target_does_not_authorize_silently_dropping_an_uncovered_nested_target() {
    let body = vec![go("top"), go("nested"), declaration("late", "int"), label("top"),
        stmt(StmtKind::If { cond: number(0), then_body: vec![label("nested")], elifs: Vec::new(), else_body: None })];
    let before = format!("{body:?}");
    assert!(prepare(&body, &HashSet::new()).is_err());
    assert_eq!(format!("{body:?}"), before);
}

#[test]
fn missing_duplicate_opaque_and_colliding_legacy_targets_reject() {
    for body in [
        vec![go("absent"), declaration("late", "int"), label("top")],
        vec![go("top"), declaration("late", "int"), label("top"), label("top")],
        vec![go("top"), declaration("late", "int"), label("top"), stmt(StmtKind::Try {
            body: Vec::new(), catches: Vec::new(), else_body: None, finally: None })],
        vec![go("top"), declaration("__c_goto_pc", "int"), label("top")],
    ] {
        assert!(prepare(&body, &HashSet::new()).is_err());
    }
}

#[test]
fn no_goto_remains_identical_but_does_not_hide_expression_owned_control() {
    let body = vec![declaration("record", "Record"), stmt(StmtKind::Return(Some(number(0))))];
    let result = prepare(&body, &HashSet::new()).unwrap();
    assert_eq!(result.path, Path::Unchanged);
    assert_eq!(format!("{:?}", result.body), format!("{body:?}"));
    let body = vec![stmt(StmtKind::Return(Some(vybe_ast::Expression::new(ExprKind::Lambda {
        params: Vec::new(), body: LambdaBody::Block(vec![go("hidden"), label("hidden")]),
        is_async: false, captures: Vec::new(),
    }))))];
    assert!(prepare(&body, &HashSet::new()).is_err());
}

#[test]
fn owned_unchanged_function_keeps_its_allocation_and_rejection_returns_original_body() {
    let body = vec![stmt(StmtKind::Return(Some(number(0))))];
    let allocation = body.as_ptr();
    let result = prepare_owned(body, &HashSet::new()).unwrap_or_else(|_| panic!("unchanged body rejected"));
    assert_eq!(result.body.as_ptr(), allocation, "do not clone every no-goto function AST");
    let body = vec![go("missing")];
    let original = format!("{body:?}");
    let allocation = body.as_ptr();
    let rejection = match prepare_owned(body, &HashSet::new()) { Err(error) => error, Ok(_) => panic!("missing target accepted") };
    assert_eq!(rejection.body.as_ptr(), allocation);
    assert_eq!(format!("{:?}", rejection.body), original);
}
