use vybe_ast::*;

fn fixture() -> Expression {
    Expression::new(ExprKind::Map(vec![
        (Expression::ident("k1"), Expression::new(ExprKind::Map(vec![
            (Expression::ident("inner_key"), Expression::ident("inner_value")),
        ]))),
        (Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident("k2a")),
            right: Box::new(Expression::ident("k2b")),
        }), Expression::ident("v2")),
    ]))
}

fn label(expression: &Expression) -> String {
    match &expression.kind {
        ExprKind::Ident(name) => name.clone(),
        ExprKind::Map(_) => "map".into(),
        ExprKind::Binary { .. } => "binary".into(),
        _ => panic!("unexpected fixture node"),
    }
}

#[test]
fn map_readers_and_normalization_visit_keys_then_values_in_postorder() {
    let mut expression = fixture();
    let original = format!("{expression:?}");
    let mut read = Vec::new();
    expression.walk_exprs(&mut |child| read.push(label(child)));
    assert_eq!(format!("{expression:?}"), original);
    let mut write = Vec::new();
    expression.walk_exprs_mut(&mut |child| write.push(label(child)));
    assert_eq!(read, write);
    assert_eq!(read, ["k1", "inner_key", "inner_value", "map", "k2a", "k2b", "binary", "v2", "map"]);
    expression.walk_exprs_mut(&mut |child| {
        if let ExprKind::Ident(name) = &mut child.kind { name.insert_str(0, "bound:"); }
    });
    let mut names = Vec::new();
    expression.walk_exprs(&mut |child| {
        if let ExprKind::Ident(name) = &child.kind { names.push(name.clone()); }
    });
    assert_eq!(names, ["bound:k1", "bound:inner_key", "bound:inner_value", "bound:k2a", "bound:k2b", "bound:v2"]);
}

#[test]
fn pruning_a_map_preserves_the_existing_boundary_callback_contract() {
    let mut expression = fixture();
    let original = format!("{expression:?}");
    let mut seen = Vec::new();
    expression.walk_exprs_mut_pruned(&mut |child| seen.push(label(child)), &mut |_| false);
    assert_eq!(seen, ["map"]);
    assert_eq!(format!("{expression:?}"), original);
    seen.clear();
    expression.walk_exprs_mut_pruned(&mut |child| seen.push(label(child)), &mut |child| {
        !matches!(&child.kind, ExprKind::Map(entries) if entries.len() == 1)
    });
    assert_eq!(seen, ["k1", "map", "k2a", "k2b", "binary", "v2", "map"]);
}

#[test]
fn duplicate_map_keys_keep_entry_order_and_source_spans() {
    let mut expression = Expression::new(ExprKind::Map(vec![
        (Expression::ident("same"), Expression::ident("first")),
        (Expression::ident("same"), Expression::ident("second")),
    ]));
    let original = format!("{expression:?}");
    let mut order = Vec::new();
    expression.walk_exprs_mut(&mut |child| order.push(label(child)));
    assert_eq!(order, ["same", "first", "same", "second", "map"]);
    assert_eq!(format!("{expression:?}"), original);
}
