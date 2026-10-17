//! Owned plain/chained name assignments with ordered contextual fact updates.
use super::*;

#[derive(Clone, Copy)]
pub(super) struct Mode {
    pub sequence: bool,
    pub contextual: bool,
}
fn target(expression: &Expression) -> bool {
    match &expression.kind {
        ExprKind::Ident(_) => true,
        ExprKind::Tuple(items) => !items.is_empty() && items.iter().all(target),
        _ => false,
    }
}
pub(super) fn finish(ast: &mut Ast<'_>, rule: Rule, raw: grameo::Span, source: &str) {
    let children = ast.frame_children(rule);
    ast.depth(false);
    if children.len() < 2 {
        return;
    }
    // Preserve the capture adapter's legacy silent-augmentation recovery,
    // including its treatment of operator-looking text in string literals.
    // Such statements keep captures until that recovery policy is migrated.
    if silent_aug_assign_op(&source[raw.start..raw.end]).is_some() {
        return;
    }
    let mut cursor = raw.start;
    for (index, id) in children.iter().enumerate() {
        let Value::Ready {
            expression,
            metadata,
            ..
        } = ast.arena.get(*id)
        else {
            return;
        };
        if metadata.raw.start < cursor || metadata.raw.end > raw.end {
            return;
        }
        let gap = source[cursor..metadata.raw.start].trim();
        if (index == 0 && !gap.is_empty()) || (index != 0 && gap != "=") {
            return;
        }
        if index + 1 < children.len() && !target(expression) {
            return;
        }
        cursor = metadata.raw.end;
    }
    if !source[cursor..raw.end]
        .trim()
        .trim_matches(';')
        .trim()
        .is_empty()
    {
        return;
    }
    let id = ast.arena.list_with_metadata(
        children,
        |values| {
            let mut expressions = Vec::with_capacity(values.len());
            let mut metadata = Vec::with_capacity(values.len());
            let mut modes = Vec::with_capacity(values.len());
            let mut replay = Vec::new();
            for value in values {
                let Value::Ready {
                    expression,
                    metadata: child,
                    replay: events,
                } = value
                else {
                    unreachable!()
                };
                modes.push(Mode {
                    sequence: child.sequence,
                    contextual: child.contextual,
                });
                metadata.push((child, events.len()));
                expressions.push(expression);
                replay.extend(events);
            }
            replay.push(Replay::List {
                count: expressions.len(),
                span: span(raw),
                tuple: true,
            });
            let value = expressions.pop().unwrap();
            (
                Value::Assignment {
                    statement: Statement::with_span(
                        StmtKind::Assign {
                            targets: expressions,
                            value,
                            by_ref: false,
                        },
                        span(raw),
                    ),
                    raw,
                    rule,
                    modes,
                    replay,
                },
                Undo::Assignment(metadata),
            )
        },
        |value, undo| {
            let (
                Value::Assignment {
                    statement,
                    mut replay,
                    ..
                },
                Undo::Assignment(metadata),
            ) = (value, undo)
            else {
                unreachable!()
            };
            let StmtKind::Assign {
                mut targets, value, ..
            } = statement.kind
            else {
                unreachable!()
            };
            targets.push(value);
            replay.pop();
            let mut events = replay.into_iter();
            targets
                .into_iter()
                .zip(metadata)
                .map(|(expression, (metadata, count))| Value::Ready {
                    expression,
                    metadata,
                    replay: events.by_ref().take(count).collect(),
                })
                .collect()
        },
    );
    ast.arena.push(id);
}
pub(super) fn normalize(
    walker: &mut PyWalker,
    statement: Statement,
    modes: &[Mode],
) -> Result<Statement, String> {
    let original_span = statement.span;
    let StmtKind::Assign {
        mut targets, value, ..
    } = statement.kind
    else {
        unreachable!()
    };
    targets.push(value);
    let count = targets.len();
    let mut operands = Vec::with_capacity(count);
    for (index, (expression, mode)) in targets.into_iter().zip(modes).enumerate() {
        let is_target = index + 1 < count;
        if is_target {
            enter_assignment_target(walker);
        }
        let expression = if mode.contextual {
            normalize_contextual_list(walker, expression, mode.sequence, false)
        } else {
            normalize_expression_list(walker, expression, mode.sequence)
        };
        if is_target {
            leave_assignment_target(walker);
        }
        operands.push(expression);
    }
    // The PEG adapter and owned path share every assignment fact/type/library
    // update and special lowering, after targets then RHS normalize in order.
    lower_python_owned_assignment(walker, None, operands)
        .map(|kind| Statement::with_span(kind, original_span))
}
