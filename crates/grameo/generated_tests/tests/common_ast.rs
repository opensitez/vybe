//! Generated reductions use shared common-AST constructors, without capture walks.
use vybe_ast::{ClassMember, ExprKind, Literal, StmtKind};
use vybe_ast_bindings::{CommonAst, Value};
#[allow(dead_code)]
mod generated {
    include!(concat!(env!("OUT_DIR"), "/common_ast.rs"));
}
#[allow(dead_code)]
mod integer {
    include!(concat!(env!("OUT_DIR"), "/common_integer.rs"));
}

#[test]
fn schema_constructor_ids_are_stable_for_a_grammar_using_only_one_constructor() {
    assert_eq!(
        integer::AstConstructor::Integer as usize,
        vybe_ast_bindings::ConstructorId::Integer as usize
    );
    let Value::Expression(expression) = integer::Parser
        .build_bound_ast::<CommonAst>(integer::Rule::integer, "123")
        .unwrap()
    else {
        panic!()
    };
    assert!(matches!(expression.kind, ExprKind::Lit(Literal::Int(123))));
}

#[test]
fn generated_bindings_construct_common_classes_functions_loops_and_returns() {
    let input = "class Box { function read() { while (1) { return 7; } return; } }\nreturn 9;";
    let Value::Module(module) = generated::Parser
        .build_bound_ast::<CommonAst>(generated::Rule::program, input)
        .unwrap()
    else {
        panic!()
    };
    assert_eq!(module.body.len(), 2);
    let StmtKind::ClassDecl { name, members, .. } = &module.body[0].kind else {
        panic!()
    };
    assert_eq!(name, "Box");
    let ClassMember::Method(method) = &members[0] else {
        panic!()
    };
    let StmtKind::FunctionDecl { name, body, .. } = &method.kind else {
        panic!()
    };
    assert_eq!(name, "read");
    let StmtKind::While {
        cond,
        body: loop_body,
        ..
    } = &body[0].kind
    else {
        panic!()
    };
    assert!(matches!(cond.kind, ExprKind::Lit(Literal::Int(1))));
    assert!(
        matches!(&loop_body[0].kind, StmtKind::Return(Some(e)) if matches!(e.kind, ExprKind::Lit(Literal::Int(7))))
    );
    assert!(matches!(body[1].kind, StmtKind::Return(None)));
    assert!(
        matches!(&module.body[1].kind, StmtKind::Return(Some(e)) if matches!(e.kind, ExprKind::Lit(Literal::Int(9))))
    );
    assert_eq!(module.body[1].span.start_line, 2);
    assert_eq!(module.body[1].span.start_col, 1);
    // Capture parsing remains available and has the exact same syntax acceptance.
    assert_eq!(
        generated::Parser
            .parse(generated::Rule::program, input)
            .unwrap()
            .consumed(),
        input.len()
    );
}

#[test]
fn generated_namespace_preserves_declaration_boundaries_and_spans() {
    let source = "namespace Outer { namespace Inner { return 1; } } return 2;";
    let module = generated::Parser
        .build_bound_ast::<CommonAst>(generated::Rule::program, source)
        .unwrap()
        .into_module();
    assert_eq!(module.body.len(), 2);
    let StmtKind::NamespaceDecl { name, body } = &module.body[0].kind else {
        panic!()
    };
    assert_eq!(name, "Outer");
    let StmtKind::NamespaceDecl { name, body } = &body[0].kind else {
        panic!()
    };
    assert_eq!(name, "Inner");
    assert!(matches!(body[0].kind, StmtKind::Return(Some(_))));
    assert_eq!(
        &source[(module.body[0].span.start_col - 1) as usize
            ..(module.body[0].span.end_col - 1) as usize],
        "namespace Outer { namespace Inner { return 1; } }"
    );
}

#[test]
fn rollback_rejects_incomplete_declarations_without_publishing_partial_ast() {
    for input in [
        "class Box {",
        "function read() { return 1;",
        "while (1) { return;",
        "return 2",
        "namespace Outer { namespace Inner { return 1; }",
    ] {
        assert!(
            generated::Parser
                .build_bound_ast::<CommonAst>(generated::Rule::program, input)
                .is_err()
        );
        assert!(
            generated::Parser
                .parse(generated::Rule::program, input)
                .is_err()
        );
    }
}
