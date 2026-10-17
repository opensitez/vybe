use vybe_ast::{ExprKind, LambdaBody, ObjectProperty, StmtKind};

#[test]
fn static_helpers_record_return_copies_fields_and_computed_source_once() {
    let module = vybe_language_c::parse(
        "struct node { int value; }; static struct node original = {7};
         static struct node snapshot(void) { return original; }
         static struct node relay(void) { return snapshot(); }
         static struct node *pointer(void) { return &original; }",
    )
    .unwrap();
    let result = |name: &str| {
        let function = module.body.iter().find(|statement| {
            matches!(&statement.kind, StmtKind::FunctionDecl { name: actual, .. } if actual == name)
        }).unwrap();
        let StmtKind::FunctionDecl { body, .. } = &function.kind else {
            panic!()
        };
        body.iter()
            .find_map(|statement| match &statement.kind {
                StmtKind::Return(Some(value)) => Some(value),
                _ => None,
            })
            .unwrap()
    };
    let ExprKind::Object(fields) = &result("snapshot").kind else {
        panic!()
    };
    assert!(fields.iter().any(|field| matches!(field,
        ObjectProperty::KeyValue { value, .. } if matches!(&value.kind,
            ExprKind::Member { object, field, .. } if field == "value"
                && matches!(&object.kind, ExprKind::Ident(name) if name == "original")))));
    let ExprKind::Call { callee, args, .. } = &result("relay").kind else {
        panic!()
    };
    let (callee, args) = if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_invoke_fixed")
    {
        (&args[0].value, &args[1..])
    } else {
        (callee.as_ref(), args.as_slice())
    };
    assert_eq!(args.len(), 1);
    assert!(matches!(&args[0].value.kind, ExprKind::Call { callee, .. }
        if matches!(&callee.kind, ExprKind::Ident(name) if name == "snapshot")));
    assert!(matches!(&callee.kind, ExprKind::Lambda {
        body: LambdaBody::Expr(body), ..
    } if matches!(&body.kind, ExprKind::Object(_))));
    assert!(!matches!(&result("pointer").kind, ExprKind::Object(_)));
}
