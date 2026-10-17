use grameo::{
    Span as Bytes,
    ast::builder::{Argument, Target},
};
use vybe_ast::class_normalize::{Access, from_method_stmt, from_method_stmt_owned};
use vybe_ast::{ExprKind, Expression, Literal, Modifiers, Span, Statement, StmtKind};
use vybe_ast_bindings::{CommonAst, ConstructorId, NormalMember, Value};
#[allow(dead_code)]
mod generated {
    include!(concat!(env!("OUT_DIR"), "/normal_class.rs"));
}

#[test]
fn generated_class_assembles_normal_members_and_recovers_failed_alternative() {
    let source = "class Box { var a = 1; static function read() { return 2; } static var b; function write() { return; } var c = 3; }";
    let Value::NormalClass(class) = generated::Parser
        .build_bound_ast::<CommonAst>(generated::Rule::program, source)
        .unwrap()
    else {
        panic!()
    };
    assert_eq!(class.name, "Box");
    assert_eq!(
        class
            .instance_fields
            .iter()
            .map(|f| f.name.as_str())
            .collect::<Vec<_>>(),
        ["a", "c"]
    );
    assert_eq!(class.static_fields[0].name, "b");
    assert!(class.static_fields[0].init.is_none());
    assert_eq!(class.instance_methods[0].source_name, "write");
    assert_eq!(class.static_methods[0].canonical_name, "read");
    assert!(class.static_methods[0].raw_modifiers.is_static);
    assert!(matches!(&class.static_methods[0].body[0].kind,
        StmtKind::Return(Some(e)) if matches!(e.kind, ExprKind::Lit(Literal::Int(2)))));
    assert_eq!(
        class.static_methods[0].span.start_col,
        source.find("static function").unwrap() as u32 + 1
    );
    assert_eq!(
        generated::Parser
            .parse(generated::Rule::program, source)
            .unwrap()
            .consumed(),
        source.len()
    );
    for invalid in [
        "class Box {",
        "class Box { var a = ; }",
        "class Box { function read() { return 2; }",
    ] {
        assert!(
            generated::Parser
                .build_bound_ast::<CommonAst>(generated::Rule::program, invalid)
                .is_err()
        );
        assert!(
            generated::Parser
                .parse(generated::Rule::program, invalid)
                .is_err()
        );
    }
}

#[test]
fn normalized_method_moves_body_allocation_and_matches_existing_normalizer() {
    let body = vec![Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Lit(Literal::Int(7)),
    ))))];
    let body_ptr = body.as_ptr();
    let reference = Statement::new(StmtKind::FunctionDecl {
        name: "read".into(),
        params: vec![],
        return_type: None,
        body: body.clone(),
        modifiers: Modifiers::default(),
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    });
    let context = CommonAst::prepare("read");
    let bytes = Bytes { start: 0, end: 4 };
    let (Value::NormalMember(NormalMember::Method { value, .. }), undo) = CommonAst::construct(
        &context,
        ConstructorId::NormalMethod as usize,
        bytes,
        vec![
            Argument::One(Value::Name("read".into())),
            Argument::One(Value::Body(body)),
            Argument::Span(bytes),
        ],
    ) else {
        panic!()
    };
    assert_eq!(value.body.as_ptr(), body_ptr);
    let expected = from_method_stmt(value.span, &reference, "read", Access::Public).unwrap();
    assert_eq!(format!("{value:?}"), format!("{expected:?}"));
    let children = CommonAst::split(
        ConstructorId::NormalMethod as usize,
        Value::NormalMember(NormalMember::Method {
            is_static: false,
            value,
        }),
        undo,
    );
    let Value::Body(body) = &children[1] else {
        panic!()
    };
    assert_eq!(body.as_ptr(), body_ptr);
}

#[test]
fn owned_method_adapter_preserves_body_allocation_and_all_existing_metadata() {
    let body = vec![Statement::new(StmtKind::Return(None))];
    let body_ptr = body.as_ptr();
    let statement = Statement::new(StmtKind::FunctionDecl {
        name: "READ".into(),
        params: vec![],
        return_type: Some("int".into()),
        body,
        modifiers: Modifiers {
            is_static: true,
            is_virtual: true,
            ..Default::default()
        },
        handles: vec![],
        is_async: true,
        is_generator: true,
        is_sub: false,
    });
    let expected = from_method_stmt(Span::default(), &statement, "read", Access::Private).unwrap();
    let owned =
        from_method_stmt_owned(Span::default(), statement, "read".into(), Access::Private).unwrap();
    assert_eq!(owned.body.as_ptr(), body_ptr);
    assert_eq!(format!("{owned:?}"), format!("{expected:?}"));
    let wrong = Statement::new(StmtKind::Return(None));
    assert!(matches!(
        from_method_stmt_owned(Span::default(), wrong, "read".into(), Access::Public)
            .unwrap_err()
            .kind,
        StmtKind::Return(None)
    ));
}
