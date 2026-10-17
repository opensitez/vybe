use vybe_ast::{ExprKind, StmtKind};

#[test]
fn trap_checks_share_one_existing_provider_without_leaking_between_scripts() {
    let source = "(module (func (export \"fail\") unreachable))\n(assert_trap (invoke \"fail\") \"unreachable\")\n(assert_trap (invoke \"fail\") \"unreachable\")";
    let module = vybe_language_wast::parse(source).unwrap();
    assert_eq!(module.body.iter().filter(|statement| matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "__wast_check_trap")).count(), 1);
    let mut calls = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |expression| {
            if let ExprKind::Call { callee, .. } = &expression.kind {
                calls += usize::from(
                    matches!(&callee.kind, ExprKind::Ident(name) if name == "__wast_check_trap"),
                );
            }
        });
    }
    assert_eq!(calls, 2);
    let next = vybe_language_wast::parse("(module)").unwrap();
    assert!(!next.body.iter().any(|statement| matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "__wast_check_trap")));
}
