use vybe_ast::*;

#[test]
fn exit_status_is_visible_to_readers_and_normalization() {
    let mut statement = Statement::new(StmtKind::Exit {
        status: Some(Expression::ident("status")),
    });
    let mut seen = Vec::new();
    statement.walk_exprs(&mut |expression| {
        if let ExprKind::Ident(name) = &expression.kind {
            seen.push(name.clone());
        }
    });
    assert_eq!(seen, ["status"]);
    statement.walk_exprs_mut(&mut |expression| {
        if matches!(&expression.kind, ExprKind::Ident(name) if name == "status") {
            *expression = Expression::int(7);
        }
    });
    assert!(matches!(statement.kind, StmtKind::Exit { status: Some(Expression {
        kind: ExprKind::Lit(Literal::Int(7)), ..
    }) }));
}

#[test]
fn immutable_analysis_reaches_nested_statement_and_expression_bodies_in_order() {
    let bigint = Expression::new(ExprKind::Lit(Literal::BigInt(7)));
    let expression = Expression::new(ExprKind::Lambda {
        params: vec![], captures: vec![], is_async: false,
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(
            Expression::new(ExprKind::Binary {
                left: Box::new(Expression::int(2)), op: BinOp::Add,
                right: Box::new(bigint),
            })
        )))]),
    });
    let statement = Statement::new(StmtKind::Block(vec![Statement::new(StmtKind::Expr(expression))]));
    let original = format!("{statement:?}");
    let mut order = Vec::new();
    statement.walk_exprs(&mut |expression| {
        order.push(match &expression.kind {
            ExprKind::Lit(Literal::Int(2)) => "int",
            ExprKind::Lit(Literal::BigInt(7)) => "bigint",
            ExprKind::Binary { .. } => "binary",
            ExprKind::Lambda { .. } => "lambda",
            _ => panic!("unexpected node"),
        });
    });
    assert_eq!(order, ["int", "bigint", "binary", "lambda"]);
    assert_eq!(format!("{statement:?}"), original);
}
