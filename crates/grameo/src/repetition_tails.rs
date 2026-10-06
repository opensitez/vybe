//! Compatibility lowering for the reference grammar compiler's repeated tails.
//!
//! A rightmost `(P ~ S)* ~ P` becomes `P ~ (S ~ P)*`. This is deliberately
//! limited to structurally equal literal/reference prefixes. The distinction
//! matters for nullable S and partial matches; ordinary PEG repetition is not
//! made non-greedy. A following expression prevents the rotation, even through grouping.
use crate::Span;
use crate::grammar::{Expr, ExprId, ExprKind, GrammarSyntax};

fn append(syntax: &mut GrammarSyntax, kind: ExprKind, span: Span) -> ExprId {
    let id = syntax.expressions.len();
    syntax.expressions.push(Expr { kind, span });
    id
}
fn union(a: Span, b: Span) -> Span {
    Span {
        start: a.start.min(b.start),
        end: a.end.max(b.end),
    }
}
fn ungroup(syntax: &GrammarSyntax, mut id: ExprId) -> ExprId {
    while let ExprKind::Group(child) = syntax.expressions[id].kind {
        id = child;
    }
    id
}
fn regroup(
    syntax: &mut GrammarSyntax,
    mut original: ExprId,
    mut replacement: ExprId,
    span: Span,
) -> ExprId {
    let mut depth = 0;
    while let ExprKind::Group(child) = syntax.expressions[original].kind {
        depth += 1;
        original = child;
    }
    for _ in 0..depth {
        replacement = append(syntax, ExprKind::Group(replacement), span);
    }
    replacement
}
fn matching_prefix(syntax: &GrammarSyntax, a: ExprId, b: ExprId) -> bool {
    let a = ungroup(syntax, a);
    let b = ungroup(syntax, b);
    match (&syntax.expressions[a].kind, &syntax.expressions[b].kind) {
        (
            ExprKind::Literal {
                text: a,
                insensitive: ai,
            },
            ExprKind::Literal {
                text: b,
                insensitive: bi,
            },
        ) => a == b && ai == bi,
        (ExprKind::Reference(a), ExprKind::Reference(b)) => a == b,
        _ => false,
    }
}
fn sequence(syntax: &GrammarSyntax, left: ExprId, right: ExprId) -> Vec<ExprId> {
    let mut work = vec![right, left];
    let mut parts = Vec::new();
    while let Some(id) = work.pop() {
        if let ExprKind::Sequence { left, right } = syntax.expressions[ungroup(syntax, id)].kind {
            work.push(right);
            work.push(left);
        } else {
            parts.push(id);
        }
    }
    parts
}
fn fold(syntax: &mut GrammarSyntax, parts: &[ExprId]) -> ExprId {
    let mut id = parts[0];
    for &right in &parts[1..] {
        let span = union(syntax.expressions[id].span, syntax.expressions[right].span);
        id = append(syntax, ExprKind::Sequence { left: id, right }, span);
    }
    id
}
fn rotate(syntax: &mut GrammarSyntax, left: ExprId, right: ExprId, span: Span) -> Option<ExprId> {
    let mut parts = sequence(syntax, left, right);
    let suffix = *parts.last()?;
    let repeated = parts[parts.len().checked_sub(2)?];
    let repetition = ungroup(syntax, repeated);
    let ExprKind::Repeat {
        expression,
        min: 0,
        max: None,
    } = syntax.expressions[repetition].kind
    else {
        return None;
    };
    let body = ungroup(syntax, expression);
    let ExprKind::Sequence { left, right } = syntax.expressions[body].kind else {
        return None;
    };
    let body_parts = sequence(syntax, left, right);
    let prefix = body_parts[0];
    if !matching_prefix(syntax, prefix, suffix) {
        return None;
    }
    let rest = fold(syntax, &body_parts[1..]);
    let loop_span = union(
        syntax.expressions[repeated].span,
        syntax.expressions[suffix].span,
    );
    let new_body = append(
        syntax,
        ExprKind::Sequence {
            left: rest,
            right: suffix,
        },
        loop_span,
    );
    let new_body = regroup(syntax, expression, new_body, loop_span);
    let new_repeat = append(
        syntax,
        ExprKind::Repeat {
            expression: new_body,
            min: 0,
            max: None,
        },
        loop_span,
    );
    let new_repeat = regroup(syntax, repeated, new_repeat, loop_span);
    parts.truncate(parts.len() - 2);
    parts.extend([prefix, new_repeat]);
    let result = fold(syntax, &parts);
    syntax.expressions[result].span = span;
    Some(result)
}

/// Append transformed parents and remap rule roots, retaining every original
/// rule/terminal ID and the arena's children-before-parent invariant. Processing
/// is iterative even for long unparenthesized sequence chains.
pub(crate) fn normalize(syntax: &mut GrammarSyntax) {
    let count = syntax.expressions.len();
    let mut interior_sequence = vec![false; count];
    for expression in &syntax.expressions {
        if let ExprKind::Sequence { left, right } = expression.kind {
            for child in [left, right] {
                let child = ungroup(syntax, child);
                if matches!(syntax.expressions[child].kind, ExprKind::Sequence { .. }) {
                    interior_sequence[child] = true;
                }
            }
        }
    }
    let mut mapped: Vec<_> = (0..count).collect();
    for id in 0..count {
        let original = syntax.expressions[id].clone();
        let kind = match original.kind {
            ExprKind::Sequence { left, right } => {
                let left = mapped[left];
                let right = mapped[right];
                if !interior_sequence[id]
                    && let Some(rotated) = rotate(syntax, left, right, original.span)
                {
                    mapped[id] = rotated;
                    continue;
                }
                ExprKind::Sequence { left, right }
            }
            ExprKind::Choice { left, right } => ExprKind::Choice {
                left: mapped[left],
                right: mapped[right],
            },
            ExprKind::Repeat {
                expression,
                min,
                max,
            } => ExprKind::Repeat {
                expression: mapped[expression],
                min,
                max,
            },
            ExprKind::Predicate {
                expression,
                positive,
            } => ExprKind::Predicate {
                expression: mapped[expression],
                positive,
            },
            ExprKind::Group(child) => ExprKind::Group(mapped[child]),
            ExprKind::Push(child) => ExprKind::Push(mapped[child]),
            other => other,
        };
        if kind != syntax.expressions[id].kind {
            mapped[id] = append(syntax, kind, original.span);
        }
    }
    for rule in &mut syntax.rules {
        rule.expression = mapped[rule.expression];
    }
}
