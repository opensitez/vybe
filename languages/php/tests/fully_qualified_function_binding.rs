use vybe_ast::{ExprKind, Literal, StmtKind};

fn assert_late_function_binding(source: &str, expected: &str) {
    let module = vybe_language_php::parse(source).expect("valid PHP");
    let statement = module.body.last().expect("call statement");
    let StmtKind::Expr(expression) = &statement.kind else { panic!("expected expression"); };
    let ExprKind::Call { callee, .. } = &expression.kind else { panic!("expected call"); };
    let ExprKind::Call { callee: resolver, args, .. } = &callee.kind else {
        panic!("late function must use the existing runtime resolver: {callee:?}");
    };
    assert!(matches!(&resolver.kind, ExprKind::Ident(name) if name == "__php_global_function"));
    assert_eq!(args.len(), 1);
    assert!(matches!(&args[0].value.kind, ExprKind::Lit(Literal::Str(name)) if name == expected));
}

#[test]
fn fully_qualified_included_function_uses_canonical_late_binding() {
    assert_late_function_binding("<?php \\LateIncluded\\answer();", "__php_fn_LateIncluded__answer");
}

#[test]
fn explicit_namespace_does_not_acquire_the_callers_namespace() {
    assert_late_function_binding("<?php namespace Caller; \\LateIncluded\\answer();", "__php_fn_LateIncluded__answer");
}

#[test]
fn explicit_global_call_has_no_current_namespace_fallback() {
    assert_late_function_binding("<?php namespace Caller; \\plain_function();", "__php_fn_plain_function");
}

#[test]
fn imported_function_alias_keeps_the_same_late_binding_contract() {
    assert_late_function_binding("<?php use function LateIncluded\\answer as imported_answer; imported_answer();", "__php_fn_LateIncluded__answer");
}
