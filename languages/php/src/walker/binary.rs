//! Shared PHP binary normalization for capture walking and direct AST bindings.
//! Receives semantic operands rather than parser-specific capture handles.
use super::*;

fn arithmetic(op: BinOp) -> bool {
    matches!(
        op,
        BinOp::Add
            | BinOp::Sub
            | BinOp::Mul
            | BinOp::Div
            | BinOp::Mod
            | BinOp::Shl
            | BinOp::Shr
            | BinOp::Concat
    )
}

/// Preserve owned operand boxes when no PHP arithmetic fold is required.
/// Both capture lowering and direct construction use these same semantics.
pub(super) fn normalize_arithmetic(expression: Expression) -> Expression {
    let ExprKind::Binary { op, left, right } = &expression.kind else {
        return expression;
    };
    let span = expression.span;
    if *op == BinOp::Div && php_expr_is_literal_zero(right) {
        return Expression::with_span(
            php_throw_named_error_expr("DivisionByZeroError", "Division by zero", &span),
            span,
        );
    }
    if matches!(op, BinOp::Add | BinOp::Sub)
        && let (Some(a), Some(b)) = (
            php_literal_or_php_int_const_i64(left),
            php_literal_or_php_int_const_i64(right),
        )
    {
        let overflow = if *op == BinOp::Add {
            a.checked_add(b).is_none()
        } else {
            a.checked_sub(b).is_none()
        };
        if overflow {
            let value = if *op == BinOp::Add {
                (a as f64) + (b as f64)
            } else {
                (a as f64) - (b as f64)
            };
            return Expression::with_span(ExprKind::Lit(Literal::Float(value)), span);
        }
    }
    expression
}

pub(super) fn reduce(
    __php_w: &mut PhpWalker,
    mut left: Expression,
    mut right: Expression,
    op: BinOp,
    right_src: &str,
    span: Span,
) -> Expression {
    if arithmetic(op) {
        return normalize_arithmetic(Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
            span,
        ));
    }
    // Resolve literal class names before choosing builtin protocols. An
    // imported global class (including an alias) is absolute even while
    // the current namespace is nonempty. Dynamic RHS values and the
    // context-dependent self/static/parent names keep their own paths.
    if op == BinOp::InstanceOf {
        if let ExprKind::Ident(name) = &right.kind {
            if !name.starts_with('$') && !matches!(name.as_str(), "self" | "static" | "parent") {
                let resolved = php_resolve_class_name(__php_w, name);
                right = Expression::with_span(ExprKind::Ident(resolved), span.clone());
            }
        }
    }
    // The shared logical operators return an operand. PHP returns a bool
    // and treats empty arrays (and "0") as false on either side.
    if matches!(op, BinOp::And | BinOp::Or | BinOp::Xor) {
        let condition = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(php_truthy_condition(left)),
                right: Box::new(php_truthy_condition(right)),
            },
            span.clone(),
        );
        // Predicates may be represented as Wasm i32 values. Materialize the
        // PHP boolean result explicitly, including when used as a value.
        left = Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(condition),
                then: Box::new(Expression::with_span(
                    ExprKind::Lit(Literal::Bool(true)),
                    span,
                )),
                else_: Box::new(Expression::with_span(
                    ExprKind::Lit(Literal::Bool(false)),
                    span,
                )),
            },
            span,
        );
        return left;
    }
    // Retain successful folds instead of repeating recursive comparisons
    // or allocating a second string merely to recover the same result.
    if matches!(op, BinOp::BitAnd | BinOp::BitOr | BinOp::BitXor) {
        if let Some(value) = php_literal_string_bitwise(&left, &right, op) {
            left = Expression::with_span(ExprKind::Lit(Literal::Str(value)), span.clone());
            return left;
        }
    }
    if matches!(op, BinOp::StrictEq | BinOp::StrictNotEq) {
        if let Some(eq) = php_literal_array_compare(&left, &right, true) {
            let value = if op == BinOp::StrictNotEq { !eq } else { eq };
            left = php_bool_for_output_context(__php_w, value, span.clone());
            return left;
        }
    }
    // Two DateTime operands compare chronologically, on their timestamps.
    //
    // The four relational operators used to be rewritten here into
    // `__php_gt`/`__php_lt`/`__php_gte`/`__php_lte` calls, which took them
    // off the shared operator path entirely. They now stay as `BinOp`s and
    // the shared emitter derives each from PHP's three-way `Compare`
    // binding (`[builtin_slots.string] compare` → `common:php.compare3`),
    // per builtinslotplan.md §2f. The old helpers coerced with
    // `parseFloat`, so `"10" < "9a"` compared 10 against 9 and answered
    // false where php answers true.
    if matches!(op, BinOp::Gt | BinOp::Lt | BinOp::GtEq | BinOp::LtEq) {
        if let Some(value) = php_fold_numeric_literal_compare(op, &left, &right) {
            left = Expression::with_span(ExprKind::Lit(Literal::Bool(value)), span.clone());
            return left;
        }
        if let (Some(left_time), Some(right_time)) = (
            php_datetime_time_expr(__php_w, &left, &span),
            php_datetime_time_expr(__php_w, &right, &span),
        ) {
            left = Expression::with_span(
                ExprKind::Binary {
                    op,
                    left: Box::new(left_time),
                    right: Box::new(right_time),
                },
                span.clone(),
            );
            return left;
        }
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
            span.clone(),
        );
    } else if matches!(op, BinOp::StrictEq | BinOp::StrictNotEq)
        && let Some(equal) = php_literal_strict_value_equal(&left, &right)
    {
        left = php_bool_for_output_context(
            __php_w,
            if op == BinOp::StrictNotEq { !equal } else { equal },
            span.clone(),
        );
    } else if matches!(op, BinOp::StrictEq | BinOp::StrictNotEq)
        && matches!((&left.kind, &right.kind), (ExprKind::Array(a), ExprKind::Array(b)) if a.is_empty() && b.is_empty())
    {
        left = php_bool_for_output_context(__php_w, op == BinOp::StrictEq, span.clone());
    } else if matches!(op, BinOp::StrictEq | BinOp::StrictNotEq)
        && (php_expr_may_be_arrayish(__php_w, &left)
            || php_expr_may_be_arrayish(__php_w, &right)
            // Unknown values can be PHP arrays even when no array fact
            // survived a call, foreach binding, or dynamic lookup. A
            // scalar literal rules out array-value equality; otherwise
            // dispatch through the existing PHP comparison adapter.
            || (!matches!(&left.kind, ExprKind::Lit(_))
                && !matches!(&right.kind, ExprKind::Lit(_))))
    {
        let eq_call = Expression::with_span(
            ExprKind::Call {
                callee: Box::new(Expression::ident("__php_strict_eq")),
                args: vec![Argument::positional(left), Argument::positional(right)],
                optional: false,
            },
            span.clone(),
        );
        left = if op == BinOp::StrictNotEq {
            Expression::with_span(
                ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(eq_call),
                },
                span.clone(),
            )
        } else {
            eq_call
        };
    } else if matches!(op, BinOp::Eq | BinOp::NotEq) {
        if let (Some(left_time), Some(right_time)) = (
            php_datetime_time_expr(__php_w, &left, &span),
            php_datetime_time_expr(__php_w, &right, &span),
        ) {
            left = Expression::with_span(
                ExprKind::Binary {
                    op: if op == BinOp::NotEq {
                        BinOp::StrictNotEq
                    } else {
                        BinOp::StrictEq
                    },
                    left: Box::new(left_time),
                    right: Box::new(right_time),
                },
                span.clone(),
            );
            return left;
        }
        if let Some(eq) = php_known_loose_object_equality(__php_w, &left, &right) {
            left = Expression::with_span(
                ExprKind::Lit(Literal::Bool(if op == BinOp::NotEq { !eq } else { eq })),
                span.clone(),
            );
            return left;
        }
        if let Some(eq) = php_literal_scalar_loose_equal(&left, &right) {
            left = Expression::with_span(
                ExprKind::Lit(Literal::Bool(if op == BinOp::NotEq { !eq } else { eq })),
                span.clone(),
            );
            return left;
        }
        if let Some(eq) = php_known_array_compare(__php_w, &left, &right, false) {
            left = php_bool_for_output_context(
                __php_w,
                if op == BinOp::NotEq { !eq } else { eq },
                span.clone(),
            );
            return left;
        }
        if php_expr_is_numeric_result(&left) && php_expr_is_numeric_result(&right) {
            left = Expression::with_span(
                ExprKind::Binary {
                    op: if op == BinOp::NotEq {
                        BinOp::StrictNotEq
                    } else {
                        BinOp::StrictEq
                    },
                    left: Box::new(left),
                    right: Box::new(right),
                },
                span.clone(),
            );
            return left;
        }
        // `==` / `!=` stays a `BinOp` so the shared emitter resolves PHP's
        // loose-equality adapter through `[builtin_slots.string] eq`.
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
            span.clone(),
        );
    } else if op == BinOp::InstanceOf
        && matches!(&right.kind, ExprKind::Ident(name)
            if name.eq_ignore_ascii_case("Closure"))
    {
        left = php_mk_call("__php_is_closure", vec![left], &span);
    } else if op == BinOp::InstanceOf
        && right_src.eq_ignore_ascii_case("self")
        && current_class_name(__php_w).is_some()
    {
        let class_name = current_class_name(__php_w).unwrap_or_default();
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(Expression::ident(&class_name)),
            },
            span.clone(),
        );
    } else if op == BinOp::InstanceOf
        && php_global_get_key_expr(&right)
            .or_else(|| php_request_global_assignment_name(&right))
            .and_then(|name| lookup_simple_string_var(__php_w, &name))
            .is_some()
    {
        let name = php_global_get_key_expr(&right)
            .or_else(|| php_request_global_assignment_name(&right))
            .unwrap_or_default();
        let class_name = lookup_simple_string_var(__php_w, &name).unwrap_or_default();
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(Expression::with_span(
                    ExprKind::Ident(php_resolve_class_name(__php_w, &class_name)),
                    span.clone(),
                )),
            },
            span.clone(),
        );
    } else if op == BinOp::InstanceOf
        && matches!(&right.kind, ExprKind::Ident(n)
            if matches!(n.as_str(), "Generator" | "Iterator" | "Traversable"))
    {
        // A vybe generator is an `ObjectKind::Continuation` with no PHP
        // `__type` stamp, so `instanceof Generator/Iterator/Traversable`
        // would be false. PHP's `Generator` implements `Iterator` (thus
        // `Traversable`), so OR in a runtime `isGenerator` predicate.
        let inst = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left.clone()),
                right: Box::new(right),
            },
            span.clone(),
        );
        let is_gen = Expression::with_span(
            ExprKind::Call {
                callee: Box::new(Expression::ident("__php_is_gen")),
                args: vec![Argument::positional(left)],
                optional: false,
            },
            span.clone(),
        );
        left = Expression::with_span(
            ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(inst),
                right: Box::new(is_gen),
            },
            span.clone(),
        );
    } else if op == BinOp::InstanceOf
        && matches!(&right.kind, ExprKind::Ident(n) if n.trim_start_matches('\\') == "ArrayObject")
    {
        left = Expression::with_span(
            ExprKind::Call {
                callee: Box::new(Expression::ident("__php_is_arrayobject")),
                args: vec![Argument::positional(left)],
                optional: false,
            },
            span.clone(),
        );
    } else if op == BinOp::InstanceOf
        && matches!(&right.kind, ExprKind::Ident(name)
            if !name.starts_with('$')
                && !matches!(name.as_str(), "self" | "static" | "parent"))
    {
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
            span.clone(),
        );
    } else if op == BinOp::Spaceship {
        if let Some(cmp) = php_literal_array_spaceship(__php_w, &left, &right) {
            left = Expression::with_span(ExprKind::Lit(Literal::Int(cmp)), span.clone());
            return left;
        }
        if let (Some((left_class, left_args)), Some((right_class, right_args))) = (
            php_new_class_and_args_from_expr(__php_w, &left),
            php_new_class_and_args_from_expr(__php_w, &right),
        ) {
            if left_class == right_class && !left_args.is_empty() && !right_args.is_empty() {
                left = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::Spaceship,
                        left: Box::new(left_args[0].clone()),
                        right: Box::new(right_args[0].clone()),
                    },
                    span.clone(),
                );
                return left;
            }
        }
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
            span.clone(),
        );
    } else {
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
            span.clone(),
        );
    }
    left
}
