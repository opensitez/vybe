use vybe_compiler::{Compiler, ast::{BinOp, ExprKind, Expression, Lang, Module, Statement, StmtKind}};
use vybe_runtime::{VM, Value};

#[test]
fn both_common_coalesce_forms_handle_missing_values_and_preserve_zero() {
    for binary in [false, true] {
        for missing in [false, true] {
            let left = if missing { Expression::ident("missing") } else { Expression::float(0.0) };
            let right = Expression::float(42.0);
            let expr = Expression::new(if binary {
                ExprKind::Binary { op: BinOp::NullCoalesce, left: Box::new(left), right: Box::new(right) }
            } else {
                ExprKind::NullCoalesce { left: Box::new(left), right: Box::new(right) }
            });
            let module = Module {
                canon: Default::default(), name: "coalesce".into(), language: Lang::Unknown,
                body: vec![Statement::new(StmtKind::Return(Some(expr)))],
                imports: vec![], directives: Default::default(),
            };
            let profile = vybe_compiler::profile::parse_profile("[info]\nname='common-test'\n[compiler]\n").unwrap();
            let chunks = Compiler::with_profile(profile).compile(&module).unwrap();
            let mut vm = VM::new();
            vybe_compiler::primitives::platforms::register_platforms(
                &mut vm, &vybe_runtime::capabilities::Capabilities::all());
            let actual = vm.run(chunks).unwrap();
            let expected = if missing { 42.0 } else { 0.0 };
            assert!(matches!(actual, Value::F64(value) if value == expected), "binary={binary}, missing={missing}: {actual:?}");
        }
    }
}
