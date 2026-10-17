//! Context-neutral owned arithmetic, lowered only in the current walker scope.
use super::*;

/// A selection hint, not a syntax recognizer. Grammar matching and the completed
/// owned-value check remain authoritative. Contextual atoms retain PEG captures.
pub(super) fn candidate(start: usize, source: &str) -> bool { candidate_for(start, source, false) }
pub(super) fn bitwise_candidate(start: usize, source: &str) -> bool { candidate_for(start, source, true) }
fn candidate_for(start: usize, source: &str, bitwise: bool) -> bool {
    let bytes = source.as_bytes();
    let mut cursor = start;
    let mut operand = true;
    let mut seen = false;
    let mut operator = false;
    let mut wide = false;
    loop {
        while bytes
            .get(cursor)
            .is_some_and(|b| matches!(b, b' ' | b'\t' | b'\r'))
        {
            cursor += 1;
        }
        let Some(&byte) = bytes.get(cursor) else {
            return seen && !operand && operator && (!bitwise || wide);
        };
        if matches!(
            byte,
            b'\n'
                | b','
                | b';'
                | b':'
                | b')'
                | b']'
                | b'}'
                | b'='
                | b'!'
        ) {
            return seen && !operand && operator && (!bitwise || wide);
        }
        if matches!(byte, b'<' | b'>' | b'|' | b'&' | b'^') {
            if !bitwise || (matches!(byte,b'<' | b'>') && bytes.get(cursor+1) != Some(&byte)) {
                return seen && !operand && operator && (!bitwise || wide);
            }
        }
        if operand {
            if byte.is_ascii_alphabetic() || byte == b'_' {
                let begin = cursor;
                cursor += 1;
                while bytes
                    .get(cursor)
                    .is_some_and(|b| b.is_ascii_alphanumeric() || *b == b'_')
                {
                    cursor += 1;
                }
                if matches!(
                    &source[begin..cursor],
                    "lambda"
                        | "not"
                        | "await"
                        | "yield"
                        | "for"
                        | "if"
                        | "else"
                        | "and"
                        | "or"
                        | "in"
                        | "is"
                ) {
                    return false;
                }
            } else if byte.is_ascii_digit()
                || (byte == b'.' && bytes.get(cursor + 1).is_some_and(u8::is_ascii_digit))
            {
                cursor += 1;
                while let Some(&next) = bytes.get(cursor) {
                    if next.is_ascii_alphanumeric() || matches!(next, b'_' | b'.') {
                        cursor += 1;
                    } else if matches!(next, b'+' | b'-')
                        && matches!(bytes.get(cursor - 1), Some(b'e' | b'E'))
                    {
                        cursor += 1;
                    } else {
                        break;
                    }
                }
            } else {
                return false;
            }
            operand = false;
            seen = true;
        } else if bitwise && matches!(byte,b'|' | b'&' | b'^' | b'<' | b'>') {
            wide = true;
            operator = true;
            cursor += if matches!(byte,b'<' | b'>') { 2 } else { 1 };
            if bytes.get(cursor) == Some(&b'=') || bytes.get(cursor) == Some(&byte) { return false; }
            operand = true;
        } else if matches!(byte, b'+' | b'-' | b'*' | b'/' | b'%' | b'@') {
            if byte == b'*' && bytes.get(cursor + 1) == Some(&b'*') {
                return false;
            }
            operator = true;
            cursor += 1;
            if byte == b'/' && bytes.get(cursor) == Some(&b'/') {
                cursor += 1;
            }
            if bytes.get(cursor) == Some(&b'=') {
                return !bitwise || wide;
            } // Grammar rejects this operator tail.
            operand = true;
        } else if byte.is_ascii_alphabetic() {
            // The additive island ends before conditional, comparison and
            // comprehension/logical keywords. The enclosing PEG rules own them.
            let end = source[cursor..]
                .find(|c: char| !c.is_ascii_alphanumeric() && c != '_')
                .map_or(source.len(), |end| cursor + end);
            return operator && (!bitwise || wide) && matches!(
                &source[cursor..end],
                "if" | "else" | "and" | "or" | "in" | "is" | "not" | "for"
            );
        } else {
            return false;
        }
    }
}
pub(super) fn begin(ast: &mut Ast<'_>, start: usize) { begin_rule(ast, Rule::additive, start); }
pub(super) fn begin_rule(ast: &mut Ast<'_>, rule: Rule, start: usize) {
    if !ast.region_candidate(rule, start) {
        return;
    }
    ast.depth(true);
    ast.arena.push(if rule == Rule::postfix { ast.member_frame } else if rule == Rule::bitor_expr { ast.bitwise_frame } else { ast.pratt_frame });
}
pub(super) fn finish(ast: &mut Ast<'_>, raw: grameo::Span, source: &str) { finish_rule(ast, Rule::additive, raw, source); }
pub(super) fn finish_rule(ast: &mut Ast<'_>, rule: Rule, raw: grameo::Span, source: &str) {
    if !ast.region_candidate(rule, raw.start) {
        return;
    }
    let children = ast.frame_children(rule);
    ast.depth(false);
    if children.len() != 1 {
        return;
    }
    let id = children[0];
    let Value::Ready { metadata, .. } = ast.arena.get(id) else {
        return;
    };
    if metadata.raw.start < raw.start
        || metadata.raw.end > raw.end
        || !source[raw.start..metadata.raw.start].trim().is_empty()
        || !source[metadata.raw.end..raw.end].trim().is_empty()
    {
        return;
    }
    ast.wrap(id, raw);
    ast.arena.edit(
        id,
        |value| {
            let Value::Ready { metadata, .. } = value else {
                unreachable!()
            };
            let previous = *metadata;
            metadata.anchor = Some(raw);
            metadata.anchor_rule = rule;
            Undo::Anchor(previous)
        },
        |value, undo| {
            let (Value::Ready { metadata, .. }, Undo::Anchor(previous)) = (value, undo) else {
                unreachable!()
            };
            *metadata = previous;
        },
    );
}
pub(super) fn complete(ast: &Ast<'_>, raw: grameo::Span, rule: Rule) -> bool {
    let values = ast.arena.values();
    if values.len() < 2 || values[values.len() - 2] != if rule == Rule::postfix { ast.member_frame } else if rule == Rule::bitor_expr { ast.bitwise_frame } else { ast.pratt_frame } {
        return false;
    }
    let Some(source) = ast.source else {
        return false;
    };
    matches!(ast.arena.get(values[values.len() - 1]), Value::Ready { metadata, .. }
        if metadata.raw.start == raw.start && metadata.raw.end <= raw.end
        && source[metadata.raw.end..raw.end].trim().is_empty())
}
pub(super) fn boundary(ast: &mut Ast<'_>, end: usize) {
    let Some(id) = ast.arena.values().last().copied() else {
        return;
    };
    if !matches!(ast.arena.get(id), Value::Ready { .. }) {
        return;
    }
    ast.arena.edit(
        id,
        |value| {
            let Value::Ready {
                expression,
                metadata,
                replay,
            } = value
            else {
                unreachable!()
            };
            let previous = Undo::Span(
                *metadata,
                expression.span,
                *replay.last_mut().unwrap().span(),
            );
            metadata.boundary = end;
            if metadata.layer == 0 {
                expression.span = span(grameo::Span {
                    start: metadata.raw.start,
                    end,
                });
                *replay.last_mut().unwrap().span() = expression.span;
            }
            previous
        },
        |value, undo| {
            let (
                Value::Ready {
                    expression,
                    metadata,
                    replay,
                },
                Undo::Span(original, old_span, event_span),
            ) = (value, undo)
            else {
                unreachable!()
            };
            *metadata = original;
            expression.span = old_span;
            *replay.last_mut().unwrap().span() = event_span;
        },
    );
}
pub(super) fn operator(text: &str) -> Option<(BinOp, u8)> {
    Some(match text.trim() {
        "+" => (BinOp::Add, 2),
        "-" => (BinOp::Sub, 2),
        "*" => (BinOp::Mul, 1),
        "/" => (BinOp::Div, 1),
        "//" => (BinOp::IDiv, 1),
        "%" => (BinOp::Mod, 1),
        "@" => (BinOp::MatMul, 1),
        "<<" => (BinOp::Shl, 3),
        ">>" => (BinOp::Shr, 3),
        "&" => (BinOp::BitAnd, 4),
        "^" => (BinOp::BitXor, 5),
        "|" => (BinOp::BitOr, 6),
        _ => return None,
    })
}
pub(super) fn reduce(ast: &mut Ast<'_>, fixity: grameo::pratt::Fixity, raw_op: grameo::Span) {
    if fixity != grameo::pratt::Fixity::InfixLeft {
        return;
    }
    let Some(source) = ast.source else {
        return;
    };
    let Some((op, layer)) = operator(&source[raw_op.start..raw_op.end]) else {
        return;
    };
    let values = ast.arena.values();
    if values.len() < 2 {
        return;
    }
    let (left, right) = (values[values.len() - 2], values[values.len() - 1]);
    let (Value::Ready { metadata: lm, .. }, Value::Ready { metadata: rm, .. }) =
        (ast.arena.get(left), ast.arena.get(right))
    else {
        return;
    };
    if lm.raw.end > raw_op.start || rm.raw.start < raw_op.end {
        return;
    }
    ast.arena.pop();
    ast.arena.pop();
    let id = ast.arena.binary_with_metadata(
        left,
        right,
        |left, right| {
            let (
                Value::Ready {
                    expression: mut left,
                    metadata: lm,
                    replay: mut le,
                },
                Value::Ready {
                    expression: mut right,
                    metadata: rm,
                    replay: re,
                },
            ) = (left, right)
            else {
                unreachable!()
            };
            let undo = Undo::Binary(lm, rm, le.len(), left.span, right.span);
            if lm.layer == layer {
                left.span = Span::default();
                *le.last_mut().unwrap().span() = left.span;
            }
            let mut re = re;
            if rm.layer == layer {
                right.span = Span::default();
                *re.last_mut().unwrap().span() = right.span;
            }
            le.extend(re);
            let raw = grameo::Span {
                start: lm.raw.start,
                end: rm.boundary.max(rm.raw.end),
            };
            let expression_span = span(raw);
            le.push(Replay::Binary {
                op,
                span: expression_span,
            });
            (
                Value::Ready {
                    expression: Expression::with_span(
                        ExprKind::Binary {
                            op,
                            left: Box::new(left),
                            right: Box::new(right),
                        },
                        expression_span,
                    ),
                    metadata: Metadata {
                        raw,
                        anchor: None,
                        anchor_rule: Rule::additive,
                        sequence: false,
                        contextual: true,
                        layer,
                        boundary: raw.end,
                    },
                    replay: le,
                },
                undo,
            )
        },
        |value, undo| {
            let (
                Value::Ready {
                    expression:
                        Expression {
                            kind: ExprKind::Binary { left, right, .. },
                            ..
                        },
                    mut replay,
                    ..
                },
                Undo::Binary(lm, rm, count, left_span, right_span),
            ) = (value, undo)
            else {
                unreachable!()
            };
            replay.pop();
            let re = replay.split_off(count);
            let (mut left, mut right) = (*left, *right);
            left.span = left_span;
            right.span = right_span;
            *replay.last_mut().unwrap().span() = left_span;
            let mut re = re;
            *re.last_mut().unwrap().span() = right_span;
            (
                Value::Ready {
                    expression: left,
                    metadata: lm,
                    replay,
                },
                Value::Ready {
                    expression: right,
                    metadata: rm,
                    replay: re,
                },
            )
        },
    );
    ast.arena.push(id);
}
pub(super) fn normalize(walker: &mut PyWalker, expression: Expression) -> Expression {
    let original_span = expression.span;
    let mut result = match expression.kind {
        ExprKind::Binary { op, left, right } => {
            let left = normalize(walker, *left);
            let right = normalize(walker, *right);
            match op {
                BinOp::Add => lower_python_owned_additive(walker, "+", left, right),
                BinOp::Sub => lower_python_owned_additive(walker, "-", left, right),
                BinOp::Mul => lower_python_owned_multiplicative(walker, "*", left, right),
                BinOp::Div => lower_python_owned_multiplicative(walker, "/", left, right),
                BinOp::IDiv => lower_python_owned_multiplicative(walker, "//", left, right),
                BinOp::Mod => lower_python_owned_multiplicative(walker, "%", left, right),
                BinOp::MatMul => lower_python_owned_multiplicative(walker, "@", left, right),
                BinOp::BitOr | BinOp::BitXor | BinOp::BitAnd => lower_python_owned_bitwise(walker, op, left, right),
                BinOp::Shl | BinOp::Shr => Expression::new(ExprKind::Binary { op, left: Box::new(left), right: Box::new(right) }),
                _ => unreachable!("certified arithmetic operator"),
            }
        }
        kind => desugar_member_reads(walker, Expression::with_span(kind, original_span)),
    };
    result.span = original_span;
    result
}
