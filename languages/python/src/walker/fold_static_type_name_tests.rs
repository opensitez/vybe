//! The original recursive expression helper remains the independent control.
use super::*;

use std::sync::atomic::{AtomicUsize, Ordering};
static ROOT_VISITS: AtomicUsize = AtomicUsize::new(0);
pub(super) fn record_visit() { ROOT_VISITS.fetch_add(1, Ordering::Relaxed); }

fn attr(class: &str) -> Expression {
    call_ident("__py_attr_read", vec![Expression::ident(class), Expression::string("__name__")])
}
fn compare(mut walker: PyWalker, body: Vec<Statement>) {
    let mut control = body.clone();
    let mut candidate = body;
    for stmt in &mut control {
        stmt.walk_exprs_mut(&mut |expr| fold_static_type_name_attr_reads(&mut walker, expr));
    }
    fold_static_type_name_attr_reads_in_stmts(&mut walker, &mut candidate);
    assert_eq!(format!("{candidate:?}"), format!("{control:?}"));
}

#[test]
fn root_only_statement_callback_preserves_nested_and_newly_enabled_frontiers() {
    let mut walker = PyWalker::default();
    walker.py_defined_classes.extend(["C".to_owned(), "__name__".to_owned()]);
    let newly_enabled = call_ident("__py_attr_read", vec![Expression::ident("C"), attr("__name__")]);
    let expression = Expression::new(ExprKind::Tuple(vec![
        attr("C"), attr("Unknown"), newly_enabled,
        Expression::new(ExprKind::Map(vec![(attr("C"), attr("Unknown"))])),
        Expression::new(ExprKind::Array(vec![ArrayElement {
            key: Some(attr("C")), value: attr("C"), spread: false, by_ref: false,
        }])),
        Expression::new(ExprKind::Ternary {
            cond: Box::new(attr("C")), then: Box::new(attr("Unknown")), else_: Box::new(attr("C")),
        }),
        Expression::new(ExprKind::Interpolation(vec![InterpolPart::Formatted(attr("C"), "s".into())])),
        Expression::new(ExprKind::Lambda {
            params: vec![], body: LambdaBody::Expr(Box::new(attr("C"))), is_async: false, captures: vec![],
        }),
    ]));
    compare(walker, vec![Statement::new(StmtKind::Expr(expression))]);
}

#[test]
fn root_rewrite_uses_current_class_sets_and_stack_without_answer_cache() {
    let mut walker = PyWalker::default();
    let mut missing = attr("C");
    fold_static_type_name_attr_read_root(&mut walker, &mut missing);
    assert!(matches!(missing.kind, ExprKind::Call { .. }));
    walker.py_defined_classes.insert("C".into());
    fold_static_type_name_attr_read_root(&mut walker, &mut missing);
    assert!(matches!(&missing.kind, ExprKind::Lit(Literal::Str(value)) if value == "C"));
    walker.py_defined_classes.remove("C");
    let mut current = attr("C");
    walker.py_current_class_stack.push("C".into());
    fold_static_type_name_attr_read_root(&mut walker, &mut current);
    assert!(matches!(&current.kind, ExprKind::Lit(Literal::Str(value)) if value == "C"));
    walker.py_current_class_stack.pop();
    let mut missing_again = attr("C");
    fold_static_type_name_attr_read_root(&mut walker, &mut missing_again);
    assert!(matches!(missing_again.kind, ExprKind::Call { .. }));
}

#[test]
fn deep_statement_tree_does_one_callback_per_actual_expression_slot() {
    let mut root = Expression::ident("x");
    for _ in 0..128 { root = Expression::new(ExprKind::Member {
        object: Box::new(root), field: "field".into(), null_safe: false,
    }); }
    let mut control = Statement::new(StmtKind::Expr(root.clone()));
    let mut candidate = Statement::new(StmtKind::Expr(root));
    let mut walker = PyWalker::default();
    let mut original_visits = 0;
    control.walk_exprs_mut(&mut |expr| counted_legacy(&mut walker, expr, &mut original_visits));
    let mut candidate_visits = 0;
    candidate.walk_exprs_mut(&mut |expr| {
        candidate_visits += 1;
        fold_static_type_name_attr_read_root(&mut walker, expr);
    });
    assert_eq!(candidate_visits, 129);
    assert_eq!(original_visits, 129 * 130 / 2);
    assert_eq!(format!("{candidate:?}"), format!("{control:?}"));
}

#[test]
fn actual_registered_classmethod_and_peg_keep_full_modules_and_errors() {
    let (mut accepted, mut rejected) = (0, 0);
    for source in [
        "class C:\n    @classmethod\n    def name(cls): return cls.__name__\nx=C.name()\n",
        "class Base:\n    @classmethod\n    def name(cls): return cls.__name__\nclass Derived(Base): pass\nx=Derived.name()\n",
        "class C:\n    @classmethod\n    def names(cls):\n        def nested(): return cls.__name__\n        return [cls.__name__, nested()]\n",
        "class C:\n    @classmethod\n    def __class_getitem__(cls, x): return (cls.__name__, x)\nx=C[int]\n",
        "class __name__:\n    @classmethod\n    def name(cls): return cls.__name__\n",
        "class C:\n    @classmethod\n    def f(cls, x=1): return f'{cls.__name__}:{x}'\n",
        "class C:\n    @classmethod\n    def f(cls): return (cls.__name__ + )\n",
    ] {
        ROOT_VISITS.store(0, Ordering::Relaxed);
        direct_ast::recipes::v51_reset_public_capture_stats();
        let actual = crate::parse(source);
        let root_visits = ROOT_VISITS.load(Ordering::Relaxed);
        let public_captures = direct_ast::recipes::v51_public_capture_count();
        if actual.is_ok() {
            accepted += 1;
            assert!(root_visits > 0, "inactive fold callback: {source}");
            assert!(public_captures > 0, "inactive registered RecipeAst: {source}");
        } else {
            rejected += 1;
        }
        println!("V51_PUBLIC accepted={} root_visits={root_visits} recipe_captures={public_captures}", actual.is_ok());
        assert_eq!(format!("{actual:?}"), format!("{:?}", crate::parse_with_peg(source)), "{source}");
    }
    assert_eq!((accepted, rejected), (6, 1));
}

fn counted_legacy(__w: &mut PyWalker, expr: &mut Expression, visits: &mut usize) {
    *visits += 1;
    if let ExprKind::Call { callee, args, .. } = &expr.kind
        && matches!(&callee.kind, ExprKind::Ident(name) if name == "__py_attr_read")
        && args.len() == 2
        && matches!(&args[1].value.kind, ExprKind::Lit(Literal::Str(field)) if field == "__name__")
        && let ExprKind::Ident(class_name) = &args[0].value.kind
        && is_defined_class(__w, class_name)
    {
        *expr = Expression::string(class_name);
        return;
    }
    match &mut expr.kind {
        ExprKind::Member { object, .. } => counted_legacy(__w, object, visits),
        ExprKind::Call { callee, args, .. } => {
            counted_legacy(__w, callee, visits);
            for arg in args {
                counted_legacy(__w, &mut arg.value, visits);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            counted_legacy(__w, left, visits);
            counted_legacy(__w, right, visits);
        }
        ExprKind::Unary { expr, .. } => counted_legacy(__w, expr, visits),
        ExprKind::Ternary { cond, then, else_ } => {
            counted_legacy(__w, cond, visits);
            counted_legacy(__w, then, visits);
            counted_legacy(__w, else_, visits);
        }
        ExprKind::Index { object, index, .. } => {
            counted_legacy(__w, object, visits);
            counted_legacy(__w, index, visits);
        }
        ExprKind::New { class, args } => {
            counted_legacy(__w, class, visits);
            for arg in args {
                counted_legacy(__w, &mut arg.value, visits);
            }
        }
        ExprKind::Assign { target, value } => {
            counted_legacy(__w, target, visits);
            counted_legacy(__w, value, visits);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    counted_legacy(__w, key, visits);
                }
                counted_legacy(__w, &mut item.value, visits);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Sequence(items) | ExprKind::Set(items) => {
            for item in items {
                counted_legacy(__w, item, visits);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        counted_legacy(__w, key, visits);
                        counted_legacy(__w, value, visits);
                    }
                    ObjectProperty::Spread(value) => counted_legacy(__w, value, visits),
                    _ => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) = part {
                    counted_legacy(__w, expr, visits);
                }
            }
        }
        _ => {}
    }
}
