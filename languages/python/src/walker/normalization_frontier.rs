//! Applicability facts collected from the already lowered common AST.
//! Collection shares the compulsory bytes/memoryview walk, after earlier rewrites.
use super::*;

#[derive(Default)]
pub(super) struct LatePassFacts {
    pub logging: bool,
    pub asyncio: bool,
    pub await_expr: bool,
    pub future: bool,
    pub enum_membership: bool,
    pub deepcopy: bool,
}
impl LatePassFacts {
    pub fn retain_all(&mut self) {
        self.logging = true;
        self.asyncio = true;
        self.await_expr = true;
        self.future = true;
    }
    pub fn observe(&mut self, expr: &Expression) {
        if matches!(expr.kind, ExprKind::Await(_)) {
            self.await_expr = true;
        }
        let ExprKind::Call { callee, args, .. } = &expr.kind else {
            return;
        };
        match &callee.kind {
            ExprKind::Member { object, field, .. } => {
                self.logging |= matches!(
                    field.as_str(),
                    "debug" | "info" | "warning" | "error" | "critical" | "exception" | "log"
                ) && args.iter().any(|arg| arg.name.as_deref() == Some("extra"));
                self.asyncio |= matches!(&object.kind, ExprKind::Ident(name) if name == "asyncio");
                self.future |= field == "add_done_callback";
            }
            ExprKind::Ident(name) => {
                self.deepcopy |= name == "__py_copy_deepcopy" && (1..=2).contains(&args.len());
                self.enum_membership |= name == "__py_contains__";
                self.asyncio |= matches!(
                    name.as_str(),
                    "run" | "create_task" | "ensure_future" | "sleep" | "gather"
                );
            }
            _ => {}
        }
    }
    pub fn retained(&self) -> usize {
        [self.logging, self.asyncio, self.await_expr, self.future]
            .into_iter()
            .filter(|flag| *flag)
            .count()
    }
}

/// The preceding syntax pass only folds literals, clones/moves existing boolean
/// operands, or creates __py_slice_new with leaf defaults. Every possible
/// constructor candidate is therefore observed by its existing postorder walk.
/// This exact shape test also covers synthetic and fallback-produced AST nodes.
#[inline]
pub(super) fn is_constructor_candidate(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, .. } => matches!(&callee.kind, ExprKind::Ident(name) if name == "new_class"),
        ExprKind::New { class, .. } => matches!(&class.kind, ExprKind::Ident(name) if matches!(name.as_str(), "__py_shlex_class" | "__py_TextWrapper")),
        _ => false,
    }
}

/// Rewrites may clone an existing initializer after postorder traversal.
pub(super) fn is_deepcopy_candidate(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Call {callee, args, ..}
        if !args.is_empty() && args.len() <= 2
        && matches!(&callee.kind, ExprKind::Ident(name) if name == "__py_copy_deepcopy"))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn constructor_frontier_survives_boolean_cloning_and_slice_operand_moves() {
        for expression in [
            Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(call_ident("new_class", vec![Expression::string("C")])),
                right: Box::new(Expression::null()),
            }),
            call_ident("slice", vec![call_ident("new_class", vec![Expression::string("C")])]),
            Expression::string("new_class __py_shlex_class __py_TextWrapper"),
        ] {
            let mut expected = Statement::new(StmtKind::Expr(expression.clone()));
            expected.walk_exprs_mut(&mut normalize_python_syntax_expr);
            expected.walk_exprs_mut(&mut |expr| normalize_python_constructor_exprs(&mut PyWalker::default(), expr));
            let mut actual = Statement::new(StmtKind::Expr(expression));
            let mut needed = false;
            actual.walk_exprs_mut(&mut |expr| {
                normalize_python_syntax_expr(expr);
                needed |= is_constructor_candidate(expr);
            });
            if needed {
                actual.walk_exprs_mut(&mut |expr| normalize_python_constructor_exprs(&mut PyWalker::default(), expr));
            }
            assert_eq!(format!("{expected:?}"), format!("{actual:?}"));
            if matches!(actual.kind, StmtKind::Expr(Expression { kind: ExprKind::Lit(_), .. })) {
                assert!(!needed, "spellings in a literal do not create semantic candidates");
            } else {
                assert!(needed, "constructor in moved or cloned operands remains visible");
            }
        }
    }

    #[test]
    fn unchanged_output_handling_does_not_retain_extra_walks() {
        let mut walker = PyWalker::default();
        let mut output = call_ident("__line", vec![Expression::int(1)]);
        let before = format!("{output:?}");
        assert_eq!(normalize_python_line_value_expr_result(&mut walker, &mut output), (true, false));
        assert!(!normalize_python_bytes_memoryview_expr_changed(&mut walker, &mut output));
        assert_eq!(format!("{output:?}"), before);
    }

    #[test]
    fn semantic_frontier_ignores_literals_and_retains_real_calls() {
        let mut facts = LatePassFacts::default();
        facts.observe(&Expression::string(
            "asyncio.sleep logging extra add_done_callback",
        ));
        assert_eq!(facts.retained(), 0);
        let mut call = call_ident("sleep", vec![]);
        facts.observe(&call);
        assert!(facts.asyncio);
        call = Expression::new(ExprKind::Await(Box::new(call)));
        facts.observe(&call);
        assert!(facts.await_expr);
        facts.observe(&call_ident("__py_contains__", vec![]));
        assert!(facts.enum_membership);
        facts.observe(&call_ident("__py_copy_deepcopy", vec![Expression::ident("x")]));
        assert!(facts.deepcopy);
        assert!(!is_deepcopy_candidate(&Expression::string(
            "__py_copy_deepcopy"
        )));
        assert!(is_deepcopy_candidate(&call_ident(
            "__py_copy_deepcopy",
            vec![Expression::ident("x")]
        )));
        facts.retain_all();
        assert_eq!(facts.retained(), 4);
    }
}
