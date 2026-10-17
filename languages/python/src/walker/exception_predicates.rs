//! Linear short-circuit guards for generated, boolean-valued type predicates.
use super::*;

fn boolean_any(mut terms: impl DoubleEndedIterator<Item = Expression>) -> Option<Expression> {
    let mut result = terms.next_back()?;
    while let Some(predicate) = terms.next_back() {
        result = Expression::new(ExprKind::Ternary {
            cond: Box::new(predicate),
            then: Box::new(Expression::bool(true)),
            else_: Box::new(result),
        });
    }
    Some(result)
}

pub(super) fn type_check(value: &Expression, target: &str) -> Option<Expression> {
    let exception_type = Expression::new(ExprKind::Index {
        object: Box::new(value.clone()),
        index: Box::new(Expression::string("__exception_type")),
        null_safe: false,
    });
    // Every term is StrictEq, hence its result is a boolean. These guards need
    // boolean short-circuit semantics, not Python's operand-valued and/or.
    // Left-nesting Or and then duplicating its normalized left operand makes
    // the generated AST exponential. A guard stores each predicate once and
    // evaluates terms in the existing exception-name order until one succeeds.
    boolean_any(py_builtin_exception_names().iter().filter_map(|candidate| {
        (py_builtin_subclass(candidate, target) == Some(true)).then(|| {
            Expression::new(ExprKind::Binary {
                op: BinOp::StrictEq,
                left: Box::new(exception_type.clone()),
                right: Box::new(Expression::string(candidate)),
            })
        })
    }))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn count(expr: &Expression) -> usize {
        let mut nodes = 0;
        expr.walk_exprs(&mut |_| nodes += 1);
        nodes
    }
    fn evaluate(expr: &Expression, actual: &str, seen: &mut Vec<String>) -> bool {
        match &expr.kind {
            ExprKind::Binary { op: BinOp::StrictEq, right, .. } => {
                let ExprKind::Lit(Literal::Str(candidate)) = &right.kind else { panic!("boolean predicate") };
                seen.push(candidate.clone());
                candidate == actual
            }
            ExprKind::Ternary { cond, then, else_ } => {
                if evaluate(cond, actual, seen) { evaluate(then, actual, seen) }
                else { evaluate(else_, actual, seen) }
            }
            ExprKind::Lit(Literal::Bool(value)) => *value,
            _ => panic!("unexpected predicate {:?}", expr.kind),
        }
    }
    #[test]
    fn exception_guards_keep_boolean_order_and_linear_normalized_size() {
        for target in ["BaseException", "Exception", "LookupError", "OSError", "UnicodeError"] {
            let expected: Vec<_> = py_builtin_exception_names().iter()
                .filter(|name| py_builtin_subclass(name, target) == Some(true)).copied().collect();
            let mut tree = type_check(&Expression::ident("error"), target).unwrap();
            let before = count(&tree);
            assert_eq!(before, expected.len() * 7 - 2);
            tree.walk_exprs_mut(&mut normalize_python_syntax_expr);
            assert_eq!(count(&tree), before);
            for actual in ["ValueError", "KeyError", "FileNotFoundError", "UnicodeDecodeError", "KeyboardInterrupt", "missing"] {
                let mut seen = Vec::new();
                assert_eq!(evaluate(&tree, actual, &mut seen), expected.contains(&actual));
                let end = expected.iter().position(|name| *name == actual).map_or(expected.len(), |index| index + 1);
                assert_eq!(seen, expected[..end]);
            }
        }
        assert!(type_check(&Expression::ident("error"), "int").is_none());
    }
    #[test]
    fn generated_guard_growth_replaces_exponential_operand_lowering() {
        let terms: Vec<_> = (0..12).map(|index| Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::ident("value")),
            right: Box::new(Expression::int(index)),
        })).collect();
        let mut old = terms.iter().cloned().reduce(|left, right| Expression::new(ExprKind::Binary {
            op: BinOp::Or, left: Box::new(left), right: Box::new(right),
        })).unwrap();
        let mut linear = boolean_any(terms.into_iter()).unwrap();
        old.walk_exprs_mut(&mut normalize_python_syntax_expr);
        linear.walk_exprs_mut(&mut normalize_python_syntax_expr);
        assert_eq!(count(&linear), 58);
        assert_eq!(count(&old), 14332);
    }
}
