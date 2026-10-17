//! Owned scalar Pratt islands with transactional contextual recapture.
use grameo::{Span, arena::EditArena, builder::Builder};
use std::collections::HashMap;
use vybe_ast::{BinOp, ExprKind, Expression, Literal, UnaryOp};

use crate::Rule;

#[derive(Debug)]
enum Value {
    Frame {
        rule: usize,
        start: usize,
    },
    Prefix(Span),
    Operand {
        expression: Option<Expression>,
        raw: Span,
        anchor: Option<Span>,
        replay: Vec<Replay>,
    },
}

#[derive(Clone, Copy, Debug)]
enum Replay {
    Integer(Span),
    Scalar(usize, Span),
    Prefix(Option<UnaryOp>),
    Binary(BinOp),
}

struct BinaryUndo {
    left: Span,
    right: Span,
    left_anchor: Option<Span>,
    right_anchor: Option<Span>,
    left_len: usize,
    folded: Option<(i64, i64)>,
}

enum Undo {
    Binary(BinaryUndo),
    Prefix {
        token: Span,
        raw: Span,
        anchor: Option<Span>,
    },
}

type Edit = (Span, Option<Span>);

pub(super) struct Cached {
    first: Option<Expression>,
    replay: Vec<Replay>,
}

impl Cached {
    pub(super) fn take(&mut self, source: &str) -> Expression {
        if let Some(first) = self.first.take() {
            return first;
        }
        let mut stack = Vec::new();
        for step in &self.replay {
            match *step {
                Replay::Integer(raw) => stack.push(integer(&source[raw.start..raw.end])),
                Replay::Scalar(rule, raw) => stack.push(scalar(rule, &source[raw.start..raw.end])),
                Replay::Prefix(op) => {
                    let child = stack.pop().unwrap();
                    stack.push(prefix(op, child));
                }
                Replay::Binary(op) => {
                    let right = stack.pop().unwrap();
                    let left = stack.pop().unwrap();
                    stack.push(combine(op, left, right));
                }
            }
        }
        assert_eq!(stack.len(), 1);
        stack.pop().unwrap()
    }
}

pub(super) struct Owned<'a> {
    source: &'a str,
    arena: EditArena<Value, Undo, Edit>,
    blocked: usize,
    side: Vec<(usize, usize)>,
    pub(super) expressions: HashMap<(usize, usize), Cached>,
}

impl<'a> Owned<'a> {
    pub(super) fn new(source: &'a str) -> Self {
        Self {
            source,
            arena: Default::default(),
            blocked: 0,
            side: Vec::new(),
            expressions: Default::default(),
        }
    }

    fn push(&mut self, value: Value) {
        let id = self.arena.alloc(value);
        self.arena.push(id);
    }

    fn active(&self) -> bool {
        self.arena.values().last().is_some_and(|id| {
            matches!(
                self.arena.get(*id),
                Value::Frame { .. } | Value::Prefix(_) | Value::Operand { anchor: None, .. }
            )
        })
    }

    fn contextual(rule: usize) -> bool {
        rule == Rule::sizeof_expression as usize
            || rule == Rule::alignof_expression as usize
            || rule == Rule::type_name as usize
            || rule == Rule::typeof_specifier as usize
            || rule == Rule::offsetof_expression as usize
            || rule == Rule::generic_expression as usize
            || rule == Rule::va_arg_expression as usize
            || rule == Rule::statement_expression as usize
    }

    fn scalar(rule: usize) -> bool {
        rule == Rule::int_literal as usize
            || rule == Rule::float_literal as usize
            || rule == Rule::char_literal as usize
            || rule == Rule::utf8_char_literal as usize
            || rule == Rule::wide_char_literal as usize
            || rule == Rule::bool_literal as usize
    }

    fn finish_prefix(&mut self, rule: usize, span: Span) {
        let values = self.arena.values();
        if values.len() < 3
            || !matches!(self.arena.get(values[values.len() - 3]),
                Value::Frame { rule: owner, start } if *owner == rule && *start == span.start)
            || !matches!(
                self.arena.get(*values.last().unwrap()),
                Value::Operand {
                    expression: Some(_),
                    ..
                }
            )
        {
            return;
        }
        let Value::Prefix(token) = self.arena.get(values[values.len() - 2]) else {
            return;
        };
        let Some(op) = prefix_kind(&self.source[token.start..token.end]) else {
            return;
        };
        let right = self.arena.pop().unwrap();
        let left = self.arena.pop().unwrap();
        let id = self.arena.binary_with_metadata(
            left,
            right,
            |left, right| {
                let Value::Prefix(token) = left else {
                    unreachable!()
                };
                let Value::Operand {
                    expression,
                    raw,
                    anchor,
                    mut replay,
                } = right
                else {
                    unreachable!()
                };
                replay.push(Replay::Prefix(op));
                (
                    Value::Operand {
                        expression: expression.map(|child| prefix(op, child)),
                        raw: Span {
                            start: token.start,
                            end: raw.end,
                        },
                        anchor: None,
                        replay,
                    },
                    Undo::Prefix { token, raw, anchor },
                )
            },
            split,
        );
        self.arena.push(id);
    }

    fn postfix(rule: usize) -> bool {
        rule == Rule::call_suffix as usize
            || rule == Rule::index_suffix as usize
            || rule == Rule::member_suffix as usize
            || rule == Rule::arrow_suffix as usize
            || rule == Rule::inc_dec_suffix as usize
    }

    fn set_blocked(&mut self, value: usize) {
        self.side.push((self.arena.checkpoint(), self.blocked));
        self.blocked = value;
    }

    fn finish_operand(&mut self, rule: usize, span: Span) {
        let values = self.arena.values();
        let complete = values.len() >= 2
            && matches!(
                self.arena.get(values[values.len()-2]),
                Value::Frame { rule: frame, start } if *frame == rule && *start == span.start
            )
            && matches!(self.arena.get(*values.last().unwrap()),
            Value::Operand { expression: Some(_), raw, .. }
            if raw.start == span.start && raw.end <= span.end);
        if complete {
            self.arena.remove_value_below_top();
            let id = *self.arena.values().last().unwrap();
            self.arena.edit(
                id,
                |value| {
                    let Value::Operand { raw, anchor, .. } = value else {
                        unreachable!()
                    };
                    let old = (*raw, *anchor);
                    *raw = span;
                    *anchor = None;
                    old
                },
                restore,
            );
            return;
        }
        while let Some(id) = self.arena.pop() {
            if matches!(self.arena.get(id), Value::Frame { rule: frame, start }
                if *frame == rule && *start == span.start)
            {
                break;
            }
        }
        self.push(Value::Operand {
            expression: None,
            raw: span,
            anchor: None,
            replay: Vec::new(),
        });
    }
}

impl Builder for Owned<'_> {
    fn supports_unobserved_failure_prefixes(&self) -> bool {
        true
    }
    fn observes_rule(&self, rule: usize) -> bool {
        Self::contextual(rule)
            || rule == Rule::binary_expression as usize
            || self.blocked == 0
                && self.active()
                && (rule == Rule::unary_expression as usize
                    || rule == Rule::prefix_op as usize
                    || Self::scalar(rule))
            || self.blocked == 0 && self.active() && Self::postfix(rule)
    }
    fn opaque_rule(&self, _: usize) -> bool {
        true
    }
    fn supports_pratt(&self, rule: usize) -> bool {
        rule == Rule::binary_expression as usize
    }
    fn supports_pratt_at(&self, rule: usize, offset: usize, source: &str) -> bool {
        self.blocked == 0
            && self.supports_pratt(rule)
            && scalar_start(&source[offset..])
            && (!source.as_bytes()[offset].is_ascii_digit()
                || numeric_operator_hint(&source[offset..]))
    }
    fn suppress_pratt_captures(&self, rule: usize) -> bool {
        self.supports_pratt(rule)
    }
    fn prune_capture_children(&self, rule: usize, span: Span) -> bool {
        self.blocked == 0
            && rule == Rule::binary_expression as usize
            && self.arena.values().last().is_some_and(|id| {
                matches!(self.arena.get(*id),
                Value::Operand { expression: Some(_), anchor: Some(anchor), .. } if *anchor == span)
            })
    }
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint() + self.side.len()
    }
    fn rollback(&mut self, mark: usize) {
        while self
            .side
            .last()
            .is_some_and(|(arena_mark, _)| arena_mark + self.side.len() - 1 >= mark)
        {
            self.blocked = self.side.pop().unwrap().1;
        }
        self.arena.rollback(mark - self.side.len());
    }
    fn begin(&mut self, rule: usize, start: usize) {
        if Self::contextual(rule) {
            self.set_blocked(self.blocked + 1);
            return;
        }
        if self.blocked > 0 {
            return;
        }
        if rule == Rule::binary_expression as usize
            && !self.active()
            && !scalar_start(&self.source[start..])
        {
            return;
        }
        if rule == Rule::binary_expression as usize
            || rule == Rule::unary_expression as usize
            || Self::postfix(rule)
        {
            self.push(Value::Frame { rule, start });
            if Self::postfix(rule) {
                // Its owner cannot represent postfix syntax. Keep descendants
                // contextual until this unsupported operand is rolled back.
                self.set_blocked(self.blocked + 1);
            }
        }
    }
    fn finish(&mut self, rule: usize, span: Span, source: &str) {
        if Self::contextual(rule) {
            self.set_blocked(self.blocked - 1);
            return;
        }
        if Self::postfix(rule) {
            self.set_blocked(self.blocked - 1);
            self.finish_operand(rule, span);
            return;
        }
        if rule == Rule::unary_expression as usize
            && self.blocked > 0
            && self
                .arena
                .values()
                .last()
                .is_some_and(|id| matches!(self.arena.get(*id), Value::Prefix(_)))
        {
            self.set_blocked(self.blocked - 1);
            self.finish_operand(rule, span);
            return;
        }
        if self.blocked > 0 {
            return;
        }
        if rule == Rule::prefix_op as usize {
            self.push(Value::Prefix(span));
            if prefix_kind(&source[span.start..span.end]).is_none() {
                self.set_blocked(self.blocked + 1);
            }
        } else if Self::scalar(rule) {
            self.push(Value::Operand {
                expression: Some(scalar(rule, &source[span.start..span.end])),
                raw: span,
                anchor: None,
                replay: vec![if rule == Rule::int_literal as usize {
                    Replay::Integer(span)
                } else {
                    Replay::Scalar(rule, span)
                }],
            });
        } else if rule == Rule::unary_expression as usize {
            self.finish_prefix(rule, span);
            self.finish_operand(rule, span);
        } else if rule == Rule::binary_expression as usize {
            let values = self.arena.values();
            if values.len() >= 2
                && matches!(self.arena.get(values[values.len()-2]),
                Value::Frame { rule: frame, start } if *frame == rule && *start == span.start)
            {
                self.arena.remove_value_below_top();
                let id = *self.arena.values().last().unwrap();
                self.arena.edit(
                    id,
                    |value| {
                        let Value::Operand { raw, anchor, .. } = value else {
                            unreachable!()
                        };
                        let old = (*raw, *anchor);
                        *anchor = Some(span);
                        old
                    },
                    restore,
                );
            } else if self.active() {
                while let Some(id) = self.arena.pop() {
                    if matches!(self.arena.get(id), Value::Frame { rule: frame, start }
                        if *frame == rule && *start == span.start)
                    {
                        break;
                    }
                }
                self.push(Value::Operand {
                    expression: None,
                    raw: span,
                    anchor: Some(span),
                    replay: Vec::new(),
                });
            }
        }
    }
    fn accepts_pratt(&self, rule: usize, span: Span) -> bool {
        let values = self.arena.values();
        // Only the pending owner's single complete operand may replace captures.
        values.len() >= 2
            && matches!(self.arena.get(values[values.len() - 2]),
                Value::Frame { rule: frame, start } if *frame == rule && *start == span.start)
            && matches!(self.arena.get(*values.last().unwrap()),
                Value::Operand { expression: Some(_), raw, .. }
                if raw.start == span.start && raw.end <= span.end
                    && self.source[raw.end..span.end].trim().is_empty())
    }
    fn reduce(&mut self, _: usize, _: grameo::pratt::Fixity, operator: Span) {
        let values = self.arena.values();
        if values.len() < 2
            || !values[values.len() - 2..].iter().all(|id| {
                matches!(
                    self.arena.get(*id),
                    Value::Operand {
                        expression: Some(_),
                        ..
                    }
                )
            })
        {
            return;
        }
        let op = operator_kind(&self.source[operator.start..operator.end]);
        let right = self.arena.pop().unwrap();
        let left = self.arena.pop().unwrap();
        let id = self.arena.binary_with_metadata(
            left,
            right,
            |left, right| {
                let Value::Operand {
                    expression: le,
                    raw: lr,
                    anchor: la,
                    replay: mut replay,
                } = left
                else {
                    unreachable!()
                };
                let Value::Operand {
                    expression: re,
                    raw: rr,
                    anchor: ra,
                    replay: right_replay,
                } = right
                else {
                    unreachable!()
                };
                let folded = match (&le, &re) {
                    (
                        Some(Expression {
                            kind: ExprKind::Lit(Literal::Int(l)),
                            ..
                        }),
                        Some(Expression {
                            kind: ExprKind::Lit(Literal::Int(r)),
                            ..
                        }),
                    ) if will_fold(op, *l, *r) => Some((*l, *r)),
                    _ => None,
                };
                let left_len = replay.len();
                replay.extend(right_replay);
                replay.push(Replay::Binary(op));
                let expression = le.zip(re).map(|(left, right)| combine(op, left, right));
                (
                    Value::Operand {
                        expression,
                        raw: Span {
                            start: lr.start,
                            end: rr.end,
                        },
                        anchor: None,
                        replay,
                    },
                    Undo::Binary(BinaryUndo {
                        left: lr,
                        right: rr,
                        left_anchor: la,
                        right_anchor: ra,
                        left_len,
                        folded,
                    }),
                )
            },
            split,
        );
        self.arena.push(id);
    }
    fn committed(&mut self, _: &str, _: usize) {
        self.expressions = std::mem::take(&mut self.arena)
            .into_live_nodes()
            .filter_map(|value| match value {
                Value::Operand {
                    expression: Some(expression),
                    anchor: Some(span),
                    replay,
                    ..
                } => Some((
                    (span.start, span.end),
                    Cached {
                        first: Some(expression),
                        replay,
                    },
                )),
                _ => None,
            })
            .collect();
        self.side.clear();
        debug_assert_eq!(self.blocked, 0);
    }
}

fn restore(value: &mut Value, old: Edit) {
    let Value::Operand { raw, anchor, .. } = value else {
        unreachable!()
    };
    (*raw, *anchor) = old;
}

fn split(value: Value, undo: Undo) -> (Value, Value) {
    let undo = match undo {
        Undo::Binary(undo) => undo,
        Undo::Prefix { token, raw, anchor } => {
            let Value::Operand {
                expression,
                mut replay,
                ..
            } = value
            else {
                unreachable!()
            };
            let Some(Replay::Prefix(op)) = replay.pop() else {
                unreachable!()
            };
            let expression = expression.map(|expression| match op {
                None => expression,
                Some(_) => {
                    let ExprKind::Unary { expr, .. } = expression.kind else {
                        unreachable!()
                    };
                    *expr
                }
            });
            return (
                Value::Prefix(token),
                Value::Operand {
                    expression,
                    raw,
                    anchor,
                    replay,
                },
            );
        }
    };
    let Value::Operand {
        expression,
        mut replay,
        ..
    } = value
    else {
        unreachable!()
    };
    replay.pop();
    let right_replay = replay.split_off(undo.left_len);
    let (left, right) = match (expression, undo.folded) {
        (_, Some((left, right))) => (Some(int(left)), Some(int(right))),
        (
            Some(Expression {
                kind: ExprKind::Binary { left, right, .. },
                ..
            }),
            None,
        ) => (Some(*left), Some(*right)),
        (None, None) => (None, None),
        _ => unreachable!(),
    };
    (
        Value::Operand {
            expression: left,
            raw: undo.left,
            anchor: undo.left_anchor,
            replay,
        },
        Value::Operand {
            expression: right,
            raw: undo.right,
            anchor: undo.right_anchor,
            replay: right_replay,
        },
    )
}

fn int(value: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Int(value)))
}

fn prefix_kind(raw: &str) -> Option<Option<UnaryOp>> {
    match raw {
        "+" => Some(None),
        "-" => Some(Some(UnaryOp::Neg)),
        "!" => Some(Some(UnaryOp::Not)),
        "~" => Some(Some(UnaryOp::BitNot)),
        _ => None,
    }
}

pub(super) fn prefix(op: Option<UnaryOp>, child: Expression) -> Expression {
    match op {
        None => child,
        Some(op) => Expression::new(ExprKind::Unary {
            op,
            expr: Box::new(child),
        }),
    }
}

fn scalar_start(source: &str) -> bool {
    source.as_bytes().first().is_some_and(|byte| {
        byte.is_ascii_digit() || matches!(byte, b'.' | b'\'' | b'+' | b'-' | b'!' | b'~')
    }) || ["L'", "u'", "U'", "u8'"]
        .iter()
        .any(|start| source.starts_with(start))
        || ["true", "false"].iter().any(|word| {
            source.strip_prefix(word).is_some_and(|tail| {
                tail.as_bytes()
                    .first()
                    .is_none_or(|byte| !byte.is_ascii_alphanumeric() && *byte != b'_')
            })
        })
}

pub(super) fn scalar(rule: usize, raw: &str) -> Expression {
    if rule == Rule::int_literal as usize {
        integer(raw)
    } else if rule == Rule::float_literal as usize {
        Expression::new(ExprKind::Lit(Literal::Float(float(raw))))
    } else if rule == Rule::bool_literal as usize {
        Expression::new(ExprKind::Lit(Literal::Bool(raw == "true")))
    } else {
        int(character(raw) as i64)
    }
}

fn float(raw: &str) -> f64 {
    let text = raw.trim().trim_end_matches(['f', 'F', 'l', 'L']);
    let Some(body) = text.strip_prefix("0x").or_else(|| text.strip_prefix("0X")) else {
        return text.parse::<f64>().unwrap_or(0.0);
    };
    let Some((mantissa, exponent)) = body.split_once('p').or_else(|| body.split_once('P')) else {
        return 0.0;
    };
    let (whole, fraction) = mantissa.split_once('.').unwrap_or((mantissa, ""));
    let mut value = 0.0;
    for ch in whole.chars() {
        let Some(digit) = ch.to_digit(16) else {
            return 0.0;
        };
        value = value * 16.0 + digit as f64;
    }
    let mut place = 1.0 / 16.0;
    for ch in fraction.chars() {
        let Some(digit) = ch.to_digit(16) else {
            return 0.0;
        };
        value += digit as f64 * place;
        place /= 16.0;
    }
    value * 2.0_f64.powi(exponent.parse::<i32>().unwrap_or(0))
}

fn character(raw: &str) -> u32 {
    // Preserve the contextual decoder's first-character and escape policy.
    let raw = raw
        .strip_prefix("u8")
        .or_else(|| raw.strip_prefix('L'))
        .or_else(|| raw.strip_prefix('u'))
        .or_else(|| raw.strip_prefix('U'))
        .unwrap_or(raw);
    let mut chars = raw[1..raw.len() - 1].chars().peekable();
    let Some(ch) = chars.next() else { return 0 };
    if ch != '\\' {
        return ch as u32;
    }
    match chars.next().unwrap_or('\0') {
        'n' => 10,
        't' => 9,
        'r' => 13,
        '0' => 0,
        'a' => 7,
        'b' => 8,
        'f' => 12,
        'v' => 11,
        '\\' => 92,
        '\'' => 39,
        '"' => 34,
        'x' => {
            let mut value = 0u32;
            while let Some(&ch) = chars.peek() {
                if !ch.is_ascii_hexdigit() {
                    break;
                }
                chars.next();
                value = value * 16 + ch.to_digit(16).unwrap_or(0);
            }
            value & 0xff
        }
        escape @ ('u' | 'U') => {
            let mut value = 0u32;
            for _ in 0..if escape == 'u' { 4 } else { 8 } {
                match chars.peek() {
                    Some(&ch) if ch.is_ascii_hexdigit() => {
                        chars.next();
                        value = value * 16 + ch.to_digit(16).unwrap_or(0);
                    }
                    _ => break,
                }
            }
            value
        }
        ch if ch.is_ascii_digit() => {
            let mut value = ch.to_digit(8).unwrap_or(0);
            for _ in 0..2 {
                match chars.peek() {
                    Some(&ch) if ch.is_ascii_digit() && ch < '8' => {
                        chars.next();
                        value = value * 8 + ch.to_digit(8).unwrap_or(0);
                    }
                    _ => break,
                }
            }
            value
        }
        ch => ch as u32,
    }
}

fn integer(raw: &str) -> Expression {
    // Match walk_literal, including default AST spans and its overflow policy.
    // Capture spans are cache/replay anchors, not replacement AST source spans.
    let suffix = raw
        .trim_start_matches(|ch: char| {
            ch.is_ascii_hexdigit() || matches!(ch, 'x' | 'X' | 'b' | 'B')
        })
        .to_ascii_lowercase();
    let unsigned = suffix.contains('u');
    let digits = raw.trim_end_matches(['u', 'U', 'l', 'L']);
    let value = integer_value(digits);
    let literal = int(value);
    let type_name = if suffix.contains('l') {
        Some(if unsigned { "uint64" } else { "int64" })
    } else if unsigned
        || (digits.starts_with('0') && (i32::MAX as i64 + 1..=u32::MAX as i64).contains(&value))
    {
        Some("uint32")
    } else {
        None
    };
    match type_name {
        Some(type_name) => Expression::new(ExprKind::Cast {
            expr: Box::new(literal),
            type_name: type_name.into(),
        }),
        None => literal,
    }
}

fn integer_value(raw: &str) -> i64 {
    let raw = raw.trim();
    if let Some(hex) = raw.strip_prefix("0x").or_else(|| raw.strip_prefix("0X")) {
        return i64::from_str_radix(hex, 16)
            .or_else(|_| u64::from_str_radix(hex, 16).map(|value| value as i64))
            .unwrap_or(0);
    }
    if let Some(bin) = raw.strip_prefix("0b").or_else(|| raw.strip_prefix("0B")) {
        return i64::from_str_radix(bin, 2)
            .or_else(|_| u64::from_str_radix(bin, 2).map(|value| value as i64))
            .unwrap_or(0);
    }
    if raw.len() > 1 && raw.starts_with('0') && raw.chars().all(|ch| ch.is_ascii_digit()) {
        return i64::from_str_radix(&raw[1..], 8)
            .or_else(|_| u64::from_str_radix(&raw[1..], 8).map(|value| value as i64))
            .unwrap_or(0);
    }
    raw.parse::<i64>()
        .or_else(|_| raw.parse::<u64>().map(|value| value as i64))
        .unwrap_or(0)
}

fn will_fold(op: BinOp, l: i64, r: i64) -> bool {
    matches!(op, BinOp::Eq | BinOp::NotEq) && i32::try_from(l).is_ok() && i32::try_from(r).is_ok()
        || matches!(op, BinOp::Shl | BinOp::Shr) && (32..64).contains(&r)
}

fn constant_binary(op: BinOp, left: &Expression, right: &Expression) -> Option<Expression> {
    if let (ExprKind::Lit(Literal::Int(l)), ExprKind::Lit(Literal::Int(r))) =
        (&left.kind, &right.kind)
    {
        if matches!(op, BinOp::Eq | BinOp::NotEq)
            && i32::try_from(*l).is_ok()
            && i32::try_from(*r).is_ok()
        {
            return Some(int(i64::from(if op == BinOp::Eq { l == r } else { l != r })));
        }
        if matches!(op, BinOp::Shl | BinOp::Shr) && (32..64).contains(r) {
            return Some(int(if op == BinOp::Shl {
                l.wrapping_shl(*r as u32)
            } else {
                l >> r
            }));
        }
    }
    None
}

fn combine(op: BinOp, left: Expression, right: Expression) -> Expression {
    constant_binary(op, &left, &right).unwrap_or_else(|| Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    }))
}

/// Complete the existing pure fold after ordered operand normalization. The
/// direct constructor retains its already-owned Binary boxes otherwise.
pub(super) fn finish_binary_syntax(value: Expression) -> Expression {
    let folded = match &value.kind {
        ExprKind::Binary { op, left, right } => constant_binary(*op, left, right),
        _ => None,
    };
    folded.unwrap_or(value)
}

pub(super) fn operator_kind(text: &str) -> BinOp {
    match text.trim() {
        "||" => BinOp::Or,
        "&&" => BinOp::And,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        "<=" => BinOp::LtEq,
        ">" => BinOp::Gt,
        ">=" => BinOp::GtEq,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        _ => unreachable!(),
    }
}

fn numeric_operator_hint(source: &str) -> bool {
    let bytes = source.as_bytes();
    let mut end = 0;
    while bytes
        .get(end)
        .is_some_and(|byte| byte.is_ascii_alphanumeric() || *byte == b'.')
    {
        end += 1;
    }
    loop {
        while bytes.get(end).is_some_and(u8::is_ascii_whitespace) {
            end += 1;
        }
        let tail = &source[end..];
        if let Some(comment) = tail.strip_prefix("/*") {
            let Some(close) = comment.find("*/") else {
                return false;
            };
            end += 2 + close + 2;
        } else if let Some(comment) = tail.strip_prefix("//") {
            end += 2 + comment.find(['\n', '\r']).unwrap_or(comment.len());
        } else {
            break;
        }
    }
    bytes.get(end).is_some_and(|byte| {
        matches!(
            byte,
            b'+' | b'-' | b'*' | b'/' | b'%' | b'<' | b'>' | b'=' | b'!' | b'&' | b'|' | b'^'
        )
    })
}

/// A mixed binary region retains full contextual operand captures. No live C
/// scope is consulted during matching, and no region is published before commit.
pub(super) struct Mixed {
    steps: Vec<MixedStep>,
    operands: usize,
    stack_capacity: usize,
}

#[derive(Clone, Copy)]
enum MixedStep {
    Operand,
    Reduce(BinOp),
}

/// Matcher-issued operands and reductions, not a second precedence parse.
pub(super) struct PrattRecipe {
    nodes: std::rc::Rc<Vec<PrattNode>>,
    root: usize,
    operands: usize,
}

#[derive(Clone, Copy)]
enum PrattNode {
    Atom(Option<grameo::tree::CaptureHandle>),
    Infix { op: BinOp, left: usize, right: usize },
}

/// Older generated C programs use literal-only silent operator probes. They
/// cannot feed the contextual walker because they omit its binary_op anchors.
pub(super) fn source_pratt_capture_contract(program: &impl grameo::program::Program) -> bool {
    use grameo::program::Instruction;
    let Some(spec) = program.source_pratt(Rule::binary_expression as usize) else { return false; };
    spec.atom == Rule::unary_expression as usize
        && spec.trailing_trivia == grameo::islands::TrailingTrivia::Preserve
        && !spec.operators.is_empty()
        && spec.operators.iter().all(|operator| {
            let rule = program.rule(operator.rule);
            if rule.mode != grameo::RuleMode::Silent
                || operator.fixity != grameo::pratt::Fixity::InfixLeft { return false; }
            let Instruction::Sequence { left, right } = program.instruction(rule.expression)
                else { return false; };
            matches!(program.instruction(left), Instruction::Predicate { positive: true, .. })
                && matches!(program.instruction(right), Instruction::Call(rule) if rule == Rule::binary_op as usize)
        })
}

impl PrattRecipe {
    pub(super) fn lower<'a>(
        &self,
        pair: &grameo::compat::Pair<'a, Rule>,
        mut lower: impl FnMut(grameo::compat::Pair<'a, Rule>) -> Expression,
    ) -> Option<Expression> {
        let PrattNode::Infix { op, left, right } = self.nodes[self.root] else { return None; };
        if let (PrattNode::Atom(Some(left)), PrattNode::Atom(Some(right))) =
            (self.nodes[left], self.nodes[right]) {
            let left = pair.resolve_capture(left)?;
            let right = pair.resolve_capture(right)?;
            if left.as_rule() != Rule::unary_expression || right.as_rule() != Rule::unary_expression {
                return None;
            }
            return Some(combine(op, lower(left), lower(right)));
        }
        // Validate all handles before any scope-sensitive callback. Resolution
        // is direct arena indexing, not a search or immediate-child traversal.
        let mut pending = vec![self.root];
        let mut plan = Vec::with_capacity(self.operands * 2 - 1);
        while let Some(node) = pending.pop() {
            plan.push(node);
            match self.nodes[node] {
                PrattNode::Atom(handle) => {
                    if pair.resolve_capture(handle?)?.as_rule() != Rule::unary_expression { return None; }
                }
                PrattNode::Infix { left, right, .. } => {
                    pending.push(left);
                    pending.push(right);
                }
            }
        }
        let mut values = Vec::new();
        for node in plan.into_iter().rev() {
            match self.nodes[node] {
                PrattNode::Atom(handle) => values.push(lower(pair.resolve_capture(handle.unwrap()).unwrap())),
                PrattNode::Infix { op, .. } => {
                    let right = values.pop().unwrap();
                    let left = values.pop().unwrap();
                    values.push(combine(op, left, right));
                }
            }
        }
        values.pop()
    }
    pub(super) fn operand_count(&self) -> usize { self.operands }
}

#[derive(Clone, Copy)]
struct PrattValue {
    node: usize,
    operands: usize,
    complete: bool,
}

enum PrattStack {
    Frame { rule: usize, start: usize, operand: bool },
    Value(PrattValue),
}

struct PrattSide {
    arena: usize,
    declarations: usize,
    nodes: usize,
    roots: usize,
}

/// Construct a shared immutable reduction graph at the actual Pratt callbacks.
/// Arena stack restoration and side-vector/declaration restoration are disjoint.
pub(super) struct SourcePrattBuilder<'a> {
    source: &'a str,
    declarations: Option<ContextualBuilder<'a>>,
    arena: EditArena<PrattStack, (), ()>,
    nodes: Vec<PrattNode>,
    roots: Vec<(Span, PrattValue)>,
    side: Vec<PrattSide>,
    pub(super) expressions: HashMap<(usize, usize), PrattRecipe>,
}

impl<'a> SourcePrattBuilder<'a> {
    pub(super) fn new(source: &'a str) -> Self {
        Self { source, declarations: Some(ContextualBuilder::declarations_only(source)),
            arena: Default::default(), nodes: Vec::new(), roots: Vec::new(),
            side: Vec::new(), expressions: HashMap::new() }
    }
    pub(super) fn native(source: &'a str) -> Self {
        Self { source, declarations: None, arena: Default::default(), nodes: Vec::new(),
            roots: Vec::new(), side: Vec::new(), expressions: HashMap::new() }
    }
    fn in_declaration(&self) -> bool {
        self.declarations.as_ref().is_some_and(|builder| builder.depth > 0)
    }
    fn observes_declaration(&self, rule: usize) -> bool {
        self.declarations.as_ref().is_some_and(|builder| builder.observes_rule(rule))
    }
    fn expects_operand(&self) -> bool {
        self.arena.values().iter().rev().find_map(|id| match self.arena.get(*id) {
            PrattStack::Frame { rule, .. } => Some(*rule == Rule::binary_expression as usize),
            _ => None,
        }).unwrap_or(false)
    }
    fn save_side(&mut self) {
        self.side.push(PrattSide { arena: self.arena.checkpoint(),
            declarations: self.declarations.as_ref().map_or(0, Builder::checkpoint),
            nodes: self.nodes.len(), roots: self.roots.len() });
    }
    fn push(&mut self, value: PrattStack) {
        let id = self.arena.alloc(value);
        self.arena.push(id);
    }
    fn pop_value(&mut self) -> PrattValue {
        let id = self.arena.pop().unwrap();
        let PrattStack::Value(value) = self.arena.get(id) else { unreachable!("Pratt operand required"); };
        *value
    }
    pub(super) fn into_declarations(self) -> HashMap<(usize, usize), PlainScalarDeclarator> {
        self.declarations.map_or_else(HashMap::new, |builder| builder.declarations)
    }
    #[cfg(test)]
    pub(super) fn receipt_storage(&self) -> (usize, usize) {
        (self.nodes.len(), self.nodes.capacity() * std::mem::size_of::<PrattNode>()
            + self.side.capacity() * std::mem::size_of::<PrattSide>()
            + self.roots.capacity() * std::mem::size_of::<(Span, PrattValue)>())
    }
}

impl Builder for SourcePrattBuilder<'_> {
    fn supports_choice_heads(&self) -> bool { super::choice_heads_enabled() }
    fn supports_unobserved_failure_prefixes(&self) -> bool { true }
    fn supports_pratt(&self, rule: usize) -> bool {
        // Declaration construction still consumes its original binary_op hooks;
        // Source Pratt intentionally suppresses operator semantic callbacks.
        !self.in_declaration() && rule == Rule::binary_expression as usize
    }
    fn observes_rule(&self, rule: usize) -> bool {
        !self.in_declaration()
            && (rule == Rule::binary_expression as usize
                || rule == Rule::unary_expression as usize && self.expects_operand())
            || self.observes_declaration(rule)
    }
    fn observes_capture(&self, rule: usize) -> bool {
        // No completed binary value escapes its frame. An empty stack therefore
        // permits Direct's certified initializer island to own its captures.
        !self.in_declaration() && rule == Rule::unary_expression as usize
            && !self.arena.values().is_empty()
    }
    fn opaque_rule(&self, rule: usize) -> bool {
        self.declarations.as_ref().is_none_or(|builder| builder.opaque_rule(rule))
    }
    fn checkpoint(&self) -> usize { self.arena.checkpoint() + self.side.len() }
    fn rollback(&mut self, mark: usize) {
        while self.side.last().is_some_and(|state| state.arena + self.side.len() - 1 >= mark) {
            let state = self.side.pop().unwrap();
            if let Some(builder) = self.declarations.as_mut() { builder.rollback(state.declarations); }
            self.nodes.truncate(state.nodes);
            self.roots.truncate(state.roots);
        }
        self.arena.rollback(mark - self.side.len());
    }
    fn begin(&mut self, rule: usize, start: usize) {
        if self.observes_declaration(rule) {
            self.save_side();
            self.declarations.as_mut().unwrap().begin(rule, start);
        }
        if !self.in_declaration()
            && (rule == Rule::binary_expression as usize || rule == Rule::unary_expression as usize) {
            let operand = rule == Rule::unary_expression as usize
                && self.expects_operand();
            self.push(PrattStack::Frame { rule, start, operand });
        }
    }
    fn finish(&mut self, _: usize, _: Span, _: &str) {
        unreachable!("checked receipt finish required")
    }
    fn try_finish_capture(&mut self, rule: usize, span: Span, source: &str,
        capture: Option<grameo::tree::CaptureHandle>,
    ) -> Result<(), grameo::builder::BuildError> {
        let declaration = self.observes_declaration(rule);
        if declaration || !self.in_declaration()
            && (rule == Rule::unary_expression as usize || rule == Rule::binary_expression as usize) {
            self.save_side();
        }
        if declaration { self.declarations.as_mut().unwrap().try_finish(rule, span, source)?; }
        if !self.in_declaration() {
            if rule == Rule::unary_expression as usize {
                let operand = loop {
                    let id = self.arena.pop().unwrap();
                    if let PrattStack::Frame { rule: owner, start, operand } = self.arena.get(id)
                        && *owner == rule && *start == span.start { break *operand; }
                };
                if operand {
                    let node = self.nodes.len();
                    self.nodes.push(PrattNode::Atom(capture));
                    self.push(PrattStack::Value(PrattValue { node, operands: 1, complete: capture.is_some() }));
                }
            }
            if rule == Rule::binary_expression as usize {
                let result = self.arena.pop().unwrap();
                let PrattStack::Value(value) = self.arena.get(result) else { unreachable!(); };
                let value = *value;
                let frame = self.arena.pop().unwrap();
                assert!(matches!(self.arena.get(frame), PrattStack::Frame { rule: owner, start, .. }
                    if *owner == rule && *start == span.start));
                if value.complete && value.operands > 1 { self.roots.push((span, value)); }
                // The enclosing unary capture supplies its own operand handle;
                // completed binary values are never Pratt atoms themselves.
            }
        }
        Ok(())
    }
    fn reduce(&mut self, _: usize, fixity: grameo::pratt::Fixity, span: Span) {
        assert_eq!(fixity, grameo::pratt::Fixity::InfixLeft);
        self.save_side();
        let right = self.pop_value();
        let left = self.pop_value();
        let node = self.nodes.len();
        self.nodes.push(PrattNode::Infix { op: operator_kind(&self.source[span.start..span.end]),
            left: left.node, right: right.node });
        self.push(PrattStack::Value(PrattValue { node, operands: left.operands + right.operands,
            complete: left.complete && right.complete }));
    }
    fn committed(&mut self, source: &str, end: usize) {
        if let Some(builder) = self.declarations.as_mut() { builder.committed(source, end); }
        assert!(self.arena.values().is_empty());
        let nodes = std::rc::Rc::new(std::mem::take(&mut self.nodes));
        self.expressions = self.roots.drain(..).map(|(span, value)| ((span.start, span.end), PrattRecipe {
            nodes: nodes.clone(), root: value.node, operands: value.operands,
        })).collect();
        self.arena = Default::default();
        self.side.clear();
    }
}

impl Mixed {
    fn new(operators: Vec<BinOp>) -> Self {
        let operands = operators.len() + 1;
        let mut steps = Vec::with_capacity(operands + operators.len());
        let mut pending = Vec::with_capacity(operators.len());
        let mut depth = 0;
        let mut stack_capacity = 0;
        // Resolve precedence once, after the matching journal has committed.
        // Reductions stay between the same contextual operand callbacks.
        for index in 0..operands {
            steps.push(MixedStep::Operand);
            depth += 1;
            stack_capacity = stack_capacity.max(depth);
            if let Some(&next) = operators.get(index) {
                while pending
                    .last()
                    .is_some_and(|&op| precedence(op) >= precedence(next))
                {
                    steps.push(MixedStep::Reduce(pending.pop().unwrap()));
                    depth -= 1;
                }
                pending.push(next);
            }
        }
        while let Some(op) = pending.pop() {
            steps.push(MixedStep::Reduce(op));
        }
        Self { steps, operands, stack_capacity }
    }

    pub(super) fn operand_count(&self) -> usize {
        self.operands
    }
    pub(super) fn lower<'a>(
        &self,
        pair: &grameo::compat::Pair<'a, Rule>,
        mut operand: impl FnMut(grameo::compat::Pair<'a, Rule>) -> Expression,
    ) -> Option<Expression> {
        let mut terms = pair
            .clone()
            .into_inner()
            .filter(|child| child.as_rule() == Rule::unary_expression);
        // Validate every capture before calling the scope-sensitive consumer.
        let count = terms.clone().count();
        if count != self.operand_count() {
            return None;
        }
        let mut values = Vec::with_capacity(self.stack_capacity);
        for step in &self.steps {
            match *step {
                MixedStep::Operand => values.push(operand(terms.next().unwrap())),
                MixedStep::Reduce(op) => {
                    let right = values.pop().unwrap();
                    let left = values.pop().unwrap();
                    values.push(combine(op, left, right));
                }
            }
        }
        values.pop()
    }
}

fn precedence(op: BinOp) -> u8 {
    match op {
        BinOp::Or => 1,
        BinOp::And => 2,
        BinOp::BitOr => 3,
        BinOp::BitXor => 4,
        BinOp::BitAnd => 5,
        BinOp::Eq | BinOp::NotEq => 6,
        BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq => 7,
        BinOp::Shl | BinOp::Shr => 8,
        BinOp::Add | BinOp::Sub => 9,
        BinOp::Mul | BinOp::Div | BinOp::Mod => 10,
        _ => unreachable!(),
    }
}

struct ContextualTokens<'a> {
    inline: [grameo::pratt::Token<'a>; 7],
    len: usize,
    spill: Vec<grameo::pratt::Token<'a>>,
}

impl<'a> ContextualTokens<'a> {
    fn new() -> Self {
        Self {
            inline: [grameo::pratt::Token {
                kind: grameo::pratt::TokenKind::Atom,
                span: Span { start: 0, end: 0 },
            }; 7],
            len: 0,
            spill: Vec::new(),
        }
    }
    fn push(&mut self, token: grameo::pratt::Token<'a>) {
        if self.len < self.inline.len() {
            self.inline[self.len] = token;
        } else {
            if self.spill.is_empty() {
                self.spill.reserve(self.inline.len() * 2);
                self.spill.extend_from_slice(&self.inline);
            }
            self.spill.push(token);
        }
        self.len += 1;
    }
    fn as_slice(&self) -> &[grameo::pratt::Token<'a>] {
        if self.spill.is_empty() { &self.inline[..self.len] } else { &self.spill }
    }
}

pub(super) fn construct_contextual_binary<'a>(
    pair: &grameo::compat::Pair<'a, Rule>,
    mut operand: impl FnMut(grameo::compat::Pair<'a, Rule>) -> Expression,
) -> Option<Expression> {
    use grameo::pratt::{Build, Error, Fixity, Limits, Operator, Table, Token, TokenKind};
    static OPERATORS: &[Operator<'static>] = &[
        Operator { symbol: "||", precedence: 1, fixity: Fixity::InfixLeft },
        Operator { symbol: "&&", precedence: 2, fixity: Fixity::InfixLeft },
        Operator { symbol: "|", precedence: 3, fixity: Fixity::InfixLeft },
        Operator { symbol: "^", precedence: 4, fixity: Fixity::InfixLeft },
        Operator { symbol: "&", precedence: 5, fixity: Fixity::InfixLeft },
        Operator { symbol: "==", precedence: 6, fixity: Fixity::InfixLeft },
        Operator { symbol: "!=", precedence: 6, fixity: Fixity::InfixLeft },
        Operator { symbol: "<", precedence: 7, fixity: Fixity::InfixLeft },
        Operator { symbol: "<=", precedence: 7, fixity: Fixity::InfixLeft },
        Operator { symbol: ">", precedence: 7, fixity: Fixity::InfixLeft },
        Operator { symbol: ">=", precedence: 7, fixity: Fixity::InfixLeft },
        Operator { symbol: "<<", precedence: 8, fixity: Fixity::InfixLeft },
        Operator { symbol: ">>", precedence: 8, fixity: Fixity::InfixLeft },
        Operator { symbol: "+", precedence: 9, fixity: Fixity::InfixLeft },
        Operator { symbol: "-", precedence: 9, fixity: Fixity::InfixLeft },
        Operator { symbol: "*", precedence: 10, fixity: Fixity::InfixLeft },
        Operator { symbol: "/", precedence: 10, fixity: Fixity::InfixLeft },
        Operator { symbol: "%", precedence: 10, fixity: Fixity::InfixLeft },
    ];
    static TABLE: std::sync::OnceLock<Table<'static>> = std::sync::OnceLock::new();
    if pair.as_rule() != Rule::binary_expression {
        return None;
    }
    let mut children = pair.clone().into_inner();
    let mut prefix: [Option<grameo::compat::Pair<'a, Rule>>; 4] =
        std::array::from_fn(|_| children.next());
    if prefix[3].is_none() && prefix[2].is_some() {
        let left = prefix[0].take()?;
        let operator = prefix[1].take()?;
        let right = prefix[2].take()?;
        // The fourth-child lookahead certifies the complete region before any
        // contextual callback. Move these handles instead of replaying captures.
        if left.as_rule() != Rule::unary_expression
            || operator.as_rule() != Rule::binary_op
            || right.as_rule() != Rule::unary_expression
        {
            return None;
        }
        let symbol = operator.as_str().trim();
        if !OPERATORS.iter().any(|operator| operator.symbol == symbol) {
            return None;
        }
        let left = operand(left);
        let right = operand(right);
        return Some(combine(operator_kind(symbol), left, right));
    }
    let mut tokens = ContextualTokens::new();
    let mut expects_operand = true;
    // Validate the complete token shape before invoking any C scope callback.
    for child in prefix.into_iter().flatten().chain(children) {
        let kind = match child.as_rule() {
            Rule::unary_expression if expects_operand => TokenKind::Atom,
            Rule::binary_op if !expects_operand => {
                let symbol = child.as_str().trim();
                if !OPERATORS.iter().any(|operator| operator.symbol == symbol) {
                    return None;
                }
                TokenKind::Symbol(symbol)
            }
            _ => return None,
        };
        let span = child.as_span();
        tokens.push(Token { kind, span: Span { start: span.start(), end: span.end() } });
        expects_operand = !expects_operand;
    }
    let tokens = tokens.as_slice();
    if tokens.is_empty() || expects_operand {
        return None;
    }
    struct Context<I, F> {
        operands: I,
        lower: F,
    }
    impl<'a, I, F> Build for Context<I, F>
    where
        I: Iterator<Item = grameo::compat::Pair<'a, Rule>>,
        F: FnMut(grameo::compat::Pair<'a, Rule>) -> Expression,
    {
        type Value = Expression;
        fn checkpoint(&self) -> usize { 0 }
        fn rollback(&mut self, _: usize) {}
        fn atom(&mut self, _: Span) -> Result<Expression, Error> {
            Ok((self.lower)(self.operands.next().expect("validated operand capture")))
        }
        fn infix(&mut self, id: usize, _: Span, left: Expression, right: Expression) -> Result<Expression, Error> {
            Ok(combine(operator_kind(OPERATORS[id].symbol), left, right))
        }
        fn prefix(&mut self, _: usize, _: Span, _: Expression) -> Result<Expression, Error> {
            unreachable!("binary captures contain no prefix tokens")
        }
        fn postfix(&mut self, _: usize, _: Span, _: Expression) -> Result<Expression, Error> {
            unreachable!("binary captures contain no postfix tokens")
        }
    }
    let mut builder = Context {
        operands: pair.clone().into_inner().filter(|child| child.as_rule() == Rule::unary_expression),
        lower: operand,
    };
    let limits = Limits { max_tokens: tokens.len(), max_stack: tokens.len() };
    Some(grameo::pratt::parse(
        TABLE.get_or_init(|| Table::new(OPERATORS).expect("unique C infix operators")),
        tokens, &mut builder, limits,
    ).expect("validated C binary captures form a complete infix expression"))
}

enum MixedValue {
    Frame(usize),
    Operator(BinOp),
    Region {
        span: Span,
        children: Vec<MixedValue>,
    },
}

#[derive(Default)]
pub(super) struct MixedBuilder {
    arena: EditArena<MixedValue, (), ()>,
    pub(super) expressions: HashMap<(usize, usize), Mixed>,
}

impl MixedBuilder {
    fn push(&mut self, value: MixedValue) {
        let id = self.arena.alloc(value);
        self.arena.push(id);
    }
}

impl Builder for MixedBuilder {
    fn supports_unobserved_failure_prefixes(&self) -> bool {
        true
    }
    fn observes_rule(&self, rule: usize) -> bool {
        rule == Rule::binary_expression as usize || rule == Rule::binary_op as usize
    }
    fn opaque_rule(&self, _: usize) -> bool {
        true
    }
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        self.arena.rollback(mark);
    }
    fn begin(&mut self, rule: usize, start: usize) {
        if rule == Rule::binary_expression as usize {
            self.push(MixedValue::Frame(start));
        }
    }
    fn finish(&mut self, rule: usize, span: Span, source: &str) {
        if rule == Rule::binary_op as usize {
            self.push(MixedValue::Operator(operator_kind(
                &source[span.start..span.end],
            )));
        } else if rule == Rule::binary_expression as usize {
            if self.arena.values().last().is_some_and(|id| {
                matches!(self.arena.get(*id), MixedValue::Frame(start) if *start == span.start)
            }) {
                // No operator or nested region was journaled. Keep the frame
                // reversible without allocating a child list for this singleton.
                self.arena.pop();
                return;
            }
            let mut children = Vec::new();
            let mut has_operator = false;
            while let Some(id) = self.arena.pop() {
                has_operator |= matches!(self.arena.get(id), MixedValue::Operator(_));
                children.push(id);
                if matches!(self.arena.get(id), MixedValue::Frame(start) if *start == span.start) {
                    break;
                }
            }
            // Popped nested regions remain live. A singleton needs no recipe,
            // but its nested binary regions still publish at entry commit.
            if !has_operator {
                return;
            }
            children.reverse();
            let id = self.arena.list(
                children,
                |children| MixedValue::Region { span, children },
                |value| {
                    let MixedValue::Region { children, .. } = value else {
                        unreachable!()
                    };
                    children
                },
            );
            self.arena.push(id);
        }
    }
    fn committed(&mut self, _: &str, _: usize) {
        self.expressions.clear();
        let mut pending: Vec<_> = std::mem::take(&mut self.arena).into_live_nodes().collect();
        while let Some(value) = pending.pop() {
            let MixedValue::Region { span, children } = value else {
                continue;
            };
            let mut operators = Vec::new();
            for child in children {
                match child {
                    MixedValue::Operator(op) => operators.push(op),
                    region @ MixedValue::Region { .. } => pending.push(region),
                    MixedValue::Frame(_) => {}
                }
            }
            self.expressions
                .insert((span.start, span.end), Mixed::new(operators));
        }
    }
}

/// Syntax proved by the constructor, not a C type or storage decision.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct PlainDeclaratorSyntax {
    pub declarator: Span,
    pub initializer: Option<Span>,
}

#[derive(Debug)]
pub(crate) struct PlainScalarDeclarator {
    pub declaration: vybe_ast::VarDeclarator,
    pub syntax: PlainDeclaratorSyntax,
}

/// Syntax construction only. Storage, binding registration and C conversions
/// belong to the ordered declaration consumer, not a speculative matcher.
pub(crate) fn construct_plain_declarator(
    name: String,
    init: Option<Expression>,
    syntax: PlainDeclaratorSyntax,
) -> PlainScalarDeclarator {
    PlainScalarDeclarator {
        declaration: vybe_ast::VarDeclarator {
            pattern: vybe_ast::BindingPattern::Ident(name),
            type_hint: None,
            init,
            array_bounds: None,
            with_events: false,
        },
        syntax,
    }
}

// Match only the grammar's trivia after a captured plain name. In particular,
// punctuation, declarator suffixes and unterminated comments are not trivia.
pub(super) fn plain_declarator_trivia(mut text: &str) -> bool {
    loop {
        text = text.trim_start_matches([' ', '\t', '\r', '\n']);
        if text.is_empty() {
            return true;
        }
        if let Some(tail) = text.strip_prefix("/*") {
            let Some(end) = tail.find("*/") else { return false; };
            text = &tail[end + 2..];
        } else if let Some(tail) = text.strip_prefix("//") {
            text = &tail[tail.find(['\r', '\n']).unwrap_or(tail.len())..];
        } else if let Some(tail) = text.strip_prefix("\\\r\n")
            .or_else(|| text.strip_prefix("\\\n"))
        {
            text = tail;
        } else {
            return false;
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
struct ContextualMark {
    mixed: usize,
    declarations: usize,
    depth: usize,
    arrays: usize,
}

/// Both consumers see one match. Snapshots refer only to reversible builder
/// state; C scopes are untouched until the walker consumes committed values.
pub(super) struct ContextualBuilder<'a> {
    pub(super) mixed: MixedBuilder,
    mixed_enabled: bool,
    ast: Option<grameo::ast::builder::AstBuilder<declarations::DeclarationTarget<'a>>>,
    marks: Vec<ContextualMark>,
    depth: usize,
    arrays: usize,
    pub(super) declarations: HashMap<(usize, usize), PlainScalarDeclarator>,
}

impl<'a> ContextualBuilder<'a> {
    pub(super) fn new(source: &'a str) -> Self {
        Self::with_mixed(source, true)
    }
    pub(super) fn declarations_only(source: &'a str) -> Self {
        Self::with_mixed(source, false)
    }
    fn with_mixed(source: &'a str, mixed_enabled: bool) -> Self {
        let mut builder = Self {
            mixed: MixedBuilder::default(),
            mixed_enabled,
            ast: Some(grameo::ast::builder::AstBuilder::with_context(
                declarations::BINDINGS,
                declarations::DeclarationContext::deferred(source),
            )),
            marks: Vec::new(),
            depth: 0,
            arrays: 0,
            declarations: HashMap::new(),
        };
        builder.save();
        builder
    }
    fn observes_declaration(&self, rule: usize) -> bool {
        (rule == Rule::program as usize
            || rule == Rule::normal_declaration as usize
            || self.depth > 0 && self.arrays == 0)
            && self.ast.as_ref().is_some_and(|ast| ast.observes_rule(rule))
    }
    fn save(&mut self) {
        let mark = ContextualMark {
            mixed: if self.mixed_enabled { self.mixed.checkpoint() } else { 0 },
            declarations: self.ast.as_ref().map_or(0, Builder::checkpoint),
            depth: self.depth,
            arrays: self.arrays,
        };
        if self.marks.last() != Some(&mark) {
            self.marks.push(mark);
        }
    }
}

impl Builder for ContextualBuilder<'_> {
    fn supports_unobserved_failure_prefixes(&self) -> bool {
        true
    }
    fn observes_rule(&self, rule: usize) -> bool {
        self.mixed_enabled && self.mixed.observes_rule(rule)
            || self.observes_declaration(rule)
            || self.depth > 0 && rule == Rule::initializer_list as usize
    }
    fn opaque_rule(&self, rule: usize) -> bool {
        !self.observes_declaration(rule)
    }
    fn checkpoint(&self) -> usize {
        self.marks.len() - 1
    }
    fn rollback(&mut self, mark: usize) {
        let state = self.marks[mark];
        if self.mixed_enabled {
            self.mixed.rollback(state.mixed);
        }
        self.ast.as_mut().unwrap().rollback(state.declarations);
        self.depth = state.depth;
        self.arrays = state.arrays;
        self.marks.truncate(mark + 1);
    }
    fn begin(&mut self, rule: usize, start: usize) {
        if self.mixed_enabled && self.mixed.observes_rule(rule) {
            self.mixed.begin(rule, start);
        }
        if self.observes_declaration(rule) {
            self.ast.as_mut().unwrap().begin(rule, start);
        }
        if rule == Rule::normal_declaration as usize {
            self.depth += 1;
        }
        if self.depth > 0 && rule == Rule::initializer_list as usize {
            self.arrays += 1;
        }
        self.save();
    }
    fn finish(&mut self, _: usize, _: Span, _: &str) {
        unreachable!("checked declaration finish required")
    }
    fn try_finish(
        &mut self,
        rule: usize,
        span: Span,
        source: &str,
    ) -> Result<(), grameo::builder::BuildError> {
        if self.mixed_enabled && self.mixed.observes_rule(rule) {
            self.mixed.finish(rule, span, source);
        }
        if self.observes_declaration(rule) {
            self.ast.as_mut().unwrap().try_finish(rule, span, source)?;
        }
        if rule == Rule::normal_declaration as usize {
            self.depth -= 1;
        }
        if rule == Rule::initializer_list as usize {
            self.arrays -= 1;
        }
        self.save();
        Ok(())
    }
    fn committed(&mut self, source: &str, end: usize) {
        if self.mixed_enabled {
            self.mixed.committed(source, end);
        }
        let mut ast = self.ast.take().unwrap();
        ast.committed(source, end);
        let mut pending = ast.into_value().ok().into_iter().collect::<Vec<_>>();
        while let Some(value) = pending.pop() {
            match value {
                declarations::Value::Declarator { span, declaration, syntax } => {
                    self.declarations
                        .insert((span.start, span.end), PlainScalarDeclarator { declaration, syntax });
                }
                declarations::Value::Declaration {
                    statement,
                    declarator_spans,
                    ..
                } => {
                    let vybe_ast::StmtKind::VarDecl { declarations, .. } = statement.kind else {
                        unreachable!()
                    };
                    for ((span, syntax), declaration) in declarator_spans.into_iter().zip(declarations) {
                        self.declarations
                            .insert((span.start, span.end), PlainScalarDeclarator { declaration, syntax });
                    }
                }
                declarations::Value::Unsupported(children)
                | declarations::Value::Specs { children, .. } => pending.extend(children),
                _ => {}
            }
        }
        self.marks.clear();
        self.save();
    }
}

pub(crate) mod declarations {
    use super::*;
    use grameo::ast::builder::{Argument, Target};
    use grameo::ast::{Binding, Cardinality, Field, Source};
    use vybe_ast::{BindingPattern, Statement, StmtKind, TypeHint, VarDeclKind, VarDeclarator};

    /// A snapshot supplied by the ordered C declaration/scope consumer. Matching
    /// must never discover storage, rename bindings, or update typedef tables.
    #[derive(Clone, Default)]
    pub struct DeclarationContext<'a> {
        source: Option<&'a str>,
        deferred: bool,
        pub declarations: HashMap<(usize, usize), DeclarationPolicy>,
        pub declarators: HashMap<(usize, usize), ScalarPolicy>,
    }

    impl<'a> DeclarationContext<'a> {
        pub fn new(source: &'a str) -> Self {
            Self {
                source: Some(source),
                ..Default::default()
            }
        }
        pub(super) fn deferred(source: &'a str) -> Self {
            Self {
                source: Some(source),
                deferred: true,
                ..Default::default()
            }
        }
    }

    #[derive(Clone)]
    pub struct DeclarationPolicy {
        pub kind: VarDeclKind,
        pub statement_span: vybe_ast::Span,
    }

    #[derive(Clone)]
    pub struct ScalarPolicy {
        pub emitted_name: String,
        pub type_hint: Option<TypeHint>,
        pub zero_if_missing: bool,
        pub integer_wrap: Option<UnsignedWidth>,
    }

    #[derive(Clone, Copy)]
    pub enum UnsignedWidth {
        W8,
        W16,
        W32,
        W64,
    }

    fn wrapped_integer(expression: &Expression, width: UnsignedWidth) -> Option<Expression> {
        let value = match &expression.kind {
            ExprKind::Lit(Literal::Int(value)) => *value,
            ExprKind::Cast { expr, .. } => match &expr.kind {
                ExprKind::Lit(Literal::Int(value)) => *value,
                _ => return None,
            },
            _ => return None,
        };
        Some(int(match width {
            UnsignedWidth::W8 => value as u8 as i64,
            UnsignedWidth::W16 => value as u16 as i64,
            UnsignedWidth::W32 => value as u32 as i64,
            UnsignedWidth::W64 => value,
        }))
    }

    const TEXT: &[Field] = &[Field {
        name: "text",
        source: Source::Text,
    }];
    const CHILDREN: &[Field] = &[Field {
        name: "values",
        source: Source::Children(Cardinality::Many),
    }];
    pub(crate) static BINDINGS: &[Binding] = &[
        Binding {
            rule: Rule::ident_name as usize,
            constructor: 0,
            fields: TEXT,
        },
        Binding {
            rule: Rule::int_literal as usize,
            constructor: 1,
            fields: TEXT,
        },
        Binding {
            rule: Rule::declaration_specifiers as usize,
            constructor: 2,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::direct_declarator as usize,
            constructor: 3,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::initializer as usize,
            constructor: 4,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::init_declarator as usize,
            constructor: 5,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::normal_declaration as usize,
            constructor: 6,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::float_literal as usize,
            constructor: 7,
            fields: TEXT,
        },
        Binding {
            rule: Rule::char_literal as usize,
            constructor: 8,
            fields: TEXT,
        },
        Binding {
            rule: Rule::utf8_char_literal as usize,
            constructor: 12,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::wide_char_literal as usize,
            constructor: 12,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::bool_literal as usize,
            constructor: 9,
            fields: TEXT,
        },
        Binding {
            rule: Rule::prefix_op as usize,
            constructor: 10,
            fields: TEXT,
        },
        Binding {
            rule: Rule::unary_expression as usize,
            constructor: 11,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::binary_op as usize,
            constructor: 13,
            fields: TEXT,
        },
        Binding {
            rule: Rule::binary_expression as usize,
            constructor: 14,
            fields: CHILDREN,
        },
        Binding {
            rule: Rule::program as usize,
            constructor: 15,
            fields: CHILDREN,
        },
    ];

    #[derive(Debug)]
    pub(crate) enum Value {
        Name {
            span: Span,
            text: String,
        },
        Scalar {
            span: Span,
            expression: Expression,
        },
        Prefix {
            span: Span,
            text: String,
        },
        Operator {
            span: Span,
            op: BinOp,
        },
        Specs {
            span: Span,
            children: Vec<Value>,
        },
        PlainName {
            span: Span,
            name_span: Span,
            text: String,
        },
        Initializer {
            span: Span,
            literal_span: Span,
            expression: Expression,
        },
        Declarator {
            span: Span,
            declaration: VarDeclarator,
            syntax: PlainDeclaratorSyntax,
        },
        Declaration {
            span: Span,
            statement: Statement,
            declarator_spans: Vec<(Span, PlainDeclaratorSyntax)>,
        },
        Unsupported(Vec<Value>),
    }

    pub(crate) enum Undo {
        None,
        Name(Span),
        Initializer(Span),
        Unary {
            child_span: Span,
            prefix: Option<(Span, String, Option<UnaryOp>)>,
        },
        Binary {
            spans: Vec<Span>,
            operators: Vec<(Span, BinOp)>,
            reductions: Vec<Reduction>,
        },
        Declarator {
            plain_span: Span,
            name_span: Span,
            original_name: Option<String>,
            initializer: Option<(Span, Span)>,
            original_initializer: Option<Expression>,
        },
        Declaration {
            specs: Value,
            declarators: Vec<(Span, PlainDeclaratorSyntax)>,
        },
    }

    pub(crate) struct Reduction {
        left: usize,
        right: usize,
        folded: Option<(i64, i64)>,
    }

    fn reduce_binary(
        values: &mut Vec<Option<Expression>>,
        stack: &mut Vec<usize>,
        reductions: &mut Vec<Reduction>,
        op: BinOp,
    ) {
        let right = stack.pop().unwrap();
        let left = stack.pop().unwrap();
        let le = values[left].take().unwrap();
        let re = values[right].take().unwrap();
        let folded = match (&le.kind, &re.kind) {
            (ExprKind::Lit(Literal::Int(l)), ExprKind::Lit(Literal::Int(r)))
                if will_fold(op, *l, *r) =>
            {
                Some((*l, *r))
            }
            _ => None,
        };
        stack.push(values.len());
        values.push(Some(combine(op, le, re)));
        reductions.push(Reduction {
            left,
            right,
            folded,
        });
    }

    pub(crate) struct DeclarationTarget<'a>(std::marker::PhantomData<&'a str>);
    impl<'a> Target for DeclarationTarget<'a> {
        type Value = Value;
        type Undo = Undo;
        type Context = DeclarationContext<'a>;
        fn prepare(_: &str) -> Self::Context {
            DeclarationContext::default()
        }
        fn validate(
            context: &Self::Context,
            _: usize,
            _: &[Field],
            span: Span,
            text: &str,
        ) -> Result<(), grameo::builder::BuildError> {
            if context.source.is_some_and(|source| {
                span.end <= source.len()
                    && (text.as_ptr() as usize).checked_sub(span.start)
                        == Some(source.as_ptr() as usize)
            }) {
                Ok(())
            } else {
                Err(grameo::builder::BuildError {
                    span,
                    message: "declaration policies belong to a different source input".into(),
                })
            }
        }

        fn construct(
            context: &Self::Context,
            constructor: usize,
            span: Span,
            arguments: Vec<Argument<'_, Value>>,
        ) -> (Value, Undo) {
            let mut children = Vec::new();
            let mut text = "";
            for argument in arguments {
                match argument {
                    Argument::Many(values) => children = values,
                    Argument::Text(value) => text = value,
                    _ => {}
                }
            }
            match constructor {
                15 => return (Value::Unsupported(children), Undo::None),
                0 => {
                    return (
                        Value::Name {
                            span,
                            text: text.into(),
                        },
                        Undo::None,
                    );
                }
                1 | 7 | 8 | 9 => {
                    return (
                        Value::Scalar {
                            span,
                            expression: scalar(
                                match constructor {
                                    1 => Rule::int_literal as usize,
                                    7 => Rule::float_literal as usize,
                                    8 => Rule::char_literal as usize,
                                    _ => Rule::bool_literal as usize,
                                },
                                text,
                            ),
                        },
                        Undo::None,
                    );
                }
                2 => return (Value::Specs { span, children }, Undo::None),
                13 => {
                    return (
                        Value::Operator {
                            span,
                            op: operator_kind(text),
                        },
                        Undo::None,
                    );
                }
                14 => {
                    if context.deferred && children.len() != 1 {
                        return (Value::Unsupported(children), Undo::None);
                    }
                    let supported = children.len() % 2 == 1
                        && children.iter().enumerate().all(|(index, child)| {
                            if index % 2 == 0 {
                                matches!(child, Value::Scalar { .. })
                            } else {
                                matches!(child, Value::Operator { .. })
                            }
                        })
                        && matches!(children.first(), Some(Value::Scalar { span: child, .. })
                            if child.start == span.start)
                        && matches!(children.last(), Some(Value::Scalar { span: child, .. })
                            if child.end == span.end);
                    if supported {
                        let mut spans = Vec::new();
                        let mut operators = Vec::new();
                        let mut values = Vec::new();
                        for child in children {
                            match child {
                                Value::Scalar { span, expression } => {
                                    spans.push(span);
                                    values.push(Some(expression));
                                }
                                Value::Operator { span, op } => operators.push((span, op)),
                                _ => unreachable!(),
                            }
                        }
                        let mut stack = Vec::with_capacity(spans.len());
                        let mut pending = Vec::with_capacity(operators.len());
                        let mut reductions = Vec::with_capacity(operators.len());
                        for index in 0..spans.len() {
                            stack.push(index);
                            if let Some(&(_, next)) = operators.get(index) {
                                while pending
                                    .last()
                                    .is_some_and(|&op| precedence(op) >= precedence(next))
                                {
                                    reduce_binary(
                                        &mut values,
                                        &mut stack,
                                        &mut reductions,
                                        pending.pop().unwrap(),
                                    );
                                }
                                pending.push(next);
                            }
                        }
                        while let Some(op) = pending.pop() {
                            reduce_binary(&mut values, &mut stack, &mut reductions, op);
                        }
                        let expression = values[stack.pop().unwrap()].take().unwrap();
                        return (
                            Value::Scalar { span, expression },
                            Undo::Binary {
                                spans,
                                operators,
                                reductions,
                            },
                        );
                    }
                }
                10 => {
                    return (
                        Value::Prefix {
                            span,
                            text: text.into(),
                        },
                        Undo::None,
                    );
                }
                12 if matches!(children.as_slice(), [Value::Scalar { span: child, .. }]
                    if child.end == span.end) =>
                {
                    let Value::Scalar {
                        span: child_span,
                        expression,
                    } = children.pop().unwrap()
                    else {
                        unreachable!()
                    };
                    return (
                        Value::Scalar { span, expression },
                        Undo::Unary {
                            child_span,
                            prefix: None,
                        },
                    );
                }
                11 => {
                    if context.deferred && children.len() != 1 {
                        return (Value::Unsupported(children), Undo::None);
                    }
                    let supported = match children.as_slice() {
                        [Value::Scalar { span: child, .. }] => *child == span,
                        [
                            Value::Prefix { span: token, text },
                            Value::Scalar { span: child, .. },
                        ] => {
                            token.start == span.start
                                && child.end == span.end
                                && prefix_kind(text).is_some()
                        }
                        _ => false,
                    };
                    if supported {
                        let Value::Scalar {
                            span: child_span,
                            expression,
                        } = children.pop().unwrap()
                        else {
                            unreachable!()
                        };
                        let prefix = children.pop().map(|value| {
                            let Value::Prefix { span, text } = value else {
                                unreachable!()
                            };
                            let op = prefix_kind(&text).unwrap();
                            (span, text, op)
                        });
                        let expression =
                            super::prefix(prefix.as_ref().and_then(|(_, _, op)| *op), expression);
                        return (
                            Value::Scalar { span, expression },
                            Undo::Unary { child_span, prefix },
                        );
                    }
                }
                3 if matches!(children.as_slice(), [Value::Name { span: name, .. }]
                    if *name == span || name.start == span.start && name.end <= span.end
                        && context.source.and_then(|source| source.get(name.end..span.end))
                            .is_some_and(plain_declarator_trivia)) =>
                {
                    let Value::Name {
                        span: name_span,
                        text,
                    } = children.pop().unwrap()
                    else {
                        unreachable!()
                    };
                    return (
                        Value::PlainName {
                            span,
                            name_span,
                            text,
                        },
                        Undo::Name(name_span),
                    );
                }
                4 if matches!(children.as_slice(), [Value::Scalar { span: literal, .. }]
                    if *literal == span) =>
                {
                    let Value::Scalar {
                        span: literal_span,
                        expression,
                    } = children.pop().unwrap()
                    else {
                        unreachable!()
                    };
                    return (
                        Value::Initializer {
                            span,
                            literal_span,
                            expression,
                        },
                        Undo::Initializer(literal_span),
                    );
                }
                5 => {
                    let policy = context.declarators.get(&(span.start, span.end));
                    if policy.is_some() || context.deferred {
                        let supported = matches!(children.as_slice(),
                            [Value::PlainName { span: name, .. }]
                            | [Value::PlainName { span: name, .. }, Value::Initializer { .. }]
                            if name.start == span.start);
                        let context_free = !context.deferred
                            || children.iter().all(|value| match value {
                                Value::Initializer { expression, .. } => match &expression.kind {
                                    ExprKind::Lit(_) => true,
                                    ExprKind::Cast { expr, .. } => {
                                        matches!(&expr.kind, ExprKind::Lit(_))
                                    }
                                    _ => false,
                                },
                                _ => true,
                            });
                        if supported && context_free {
                            let mut values = children.into_iter();
                            let Value::PlainName {
                                span: plain_span,
                                name_span,
                                text,
                            } = values.next().unwrap()
                            else {
                                unreachable!()
                            };
                            let (mut init, initializer) = match values.next() {
                                Some(Value::Initializer {
                                    span,
                                    literal_span,
                                    expression,
                                }) => (Some(expression), Some((span, literal_span))),
                                None => (
                                    policy.is_some_and(|p| p.zero_if_missing).then(|| int(0)),
                                    None,
                                ),
                                _ => unreachable!(),
                            };
                            let mut original_initializer = None;
                            if let (Some(value), Some(width)) =
                                (&init, policy.and_then(|p| p.integer_wrap))
                            {
                                if let Some(wrapped) = wrapped_integer(value, width) {
                                    original_initializer = init.replace(wrapped);
                                }
                            }
                            let (name, original_name) = match policy {
                                Some(policy) if text != policy.emitted_name => {
                                    (policy.emitted_name.clone(), Some(text))
                                }
                                _ => (text, None),
                            };
                            let syntax = PlainDeclaratorSyntax {
                                declarator: name_span,
                                initializer: initializer.map(|(span, _)| span),
                            };
                            let mut constructed = construct_plain_declarator(name, init, syntax);
                            constructed.declaration.type_hint = policy.and_then(|p| p.type_hint.clone());
                            return (
                                Value::Declarator {
                                    span,
                                    syntax: constructed.syntax,
                                    declaration: constructed.declaration,
                                },
                                Undo::Declarator {
                                    plain_span,
                                    name_span,
                                    original_name,
                                    initializer,
                                    original_initializer,
                                },
                            );
                        }
                    }
                }
                6 => {
                    let policy = context.declarations.get(&(span.start, span.end));
                    if policy.is_some() || context.deferred {
                        if matches!(children.first(), Some(Value::Specs { span: specs, .. }) if specs.start == span.start)
                            && children.len() > 1
                            && children[1..]
                                .iter()
                                .all(|value| matches!(value, Value::Declarator { .. }))
                        {
                            let mut values = children.into_iter();
                            let specs = values.next().unwrap();
                            let mut declarations = Vec::new();
                            let mut declarators = Vec::new();
                            for value in values {
                                let Value::Declarator { span, declaration, syntax } = value else {
                                    unreachable!()
                                };
                                declarations.push(declaration);
                                declarators.push((span, syntax));
                            }
                            let mut statement = Statement::new(StmtKind::VarDecl {
                                declarations,
                                kind: policy.map_or(VarDeclKind::Var, |p| p.kind.clone()),
                            });
                            statement.span =
                                policy.map_or(Default::default(), |p| p.statement_span);
                            return (
                                Value::Declaration {
                                    span,
                                    statement,
                                    declarator_spans: declarators.clone(),
                                },
                                Undo::Declaration { specs, declarators },
                            );
                        }
                    }
                }
                _ => {}
            }
            (Value::Unsupported(children), Undo::None)
        }

        fn split(_: usize, value: Value, undo: Undo) -> Vec<Value> {
            match (value, undo) {
                (Value::Specs { children, .. } | Value::Unsupported(children), Undo::None) => {
                    children
                }
                (Value::Name { .. } | Value::Scalar { .. } | Value::Prefix { .. }, Undo::None) => {
                    Vec::new()
                }
                (Value::Operator { .. }, Undo::None) => Vec::new(),
                (
                    Value::Scalar { expression, .. },
                    Undo::Binary {
                        spans,
                        operators,
                        reductions,
                    },
                ) => {
                    let mut values: Vec<Option<Expression>> =
                        (0..spans.len() + reductions.len()).map(|_| None).collect();
                    *values.last_mut().unwrap() = Some(expression);
                    for (index, reduction) in reductions.into_iter().enumerate().rev() {
                        let value = values[spans.len() + index].take().unwrap();
                        let (left, right) = if let Some((l, r)) = reduction.folded {
                            (int(l), int(r))
                        } else {
                            let ExprKind::Binary { left, right, .. } = value.kind else {
                                unreachable!()
                            };
                            (*left, *right)
                        };
                        values[reduction.left] = Some(left);
                        values[reduction.right] = Some(right);
                    }
                    let mut children = Vec::with_capacity(spans.len() + operators.len());
                    let mut operators = operators.into_iter();
                    for (index, span) in spans.into_iter().enumerate() {
                        children.push(Value::Scalar {
                            span,
                            expression: values[index].take().unwrap(),
                        });
                        if let Some((span, op)) = operators.next() {
                            children.push(Value::Operator { span, op });
                        }
                    }
                    children
                }
                (Value::Scalar { mut expression, .. }, Undo::Unary { child_span, prefix }) => {
                    let mut children = Vec::new();
                    if let Some((span, text, op)) = prefix {
                        if op.is_some() {
                            let ExprKind::Unary { expr, .. } = expression.kind else {
                                unreachable!()
                            };
                            expression = *expr;
                        }
                        children.push(Value::Prefix { span, text });
                    }
                    children.push(Value::Scalar {
                        span: child_span,
                        expression,
                    });
                    children
                }
                (Value::PlainName { text, .. }, Undo::Name(span)) => {
                    vec![Value::Name { span, text }]
                }
                (Value::Initializer { expression, .. }, Undo::Initializer(span)) => {
                    vec![Value::Scalar { span, expression }]
                }
                (
                    Value::Declarator {
                        mut declaration, ..
                    },
                    Undo::Declarator {
                        plain_span,
                        name_span,
                        original_name,
                        initializer,
                        original_initializer,
                    },
                ) => {
                    let BindingPattern::Ident(name) = declaration.pattern else {
                        unreachable!()
                    };
                    let mut values = vec![Value::PlainName {
                        span: plain_span,
                        name_span,
                        text: original_name.unwrap_or(name),
                    }];
                    if let Some((span, literal_span)) = initializer {
                        values.push(Value::Initializer {
                            span,
                            literal_span,
                            expression: original_initializer
                                .unwrap_or_else(|| declaration.init.take().unwrap()),
                        });
                    }
                    values
                }
                (
                    Value::Declaration { statement, .. },
                    Undo::Declaration { specs, declarators },
                ) => {
                    let StmtKind::VarDecl { declarations, .. } = statement.kind else {
                        unreachable!()
                    };
                    let mut values = vec![specs];
                    values.extend(
                        declarators
                            .into_iter()
                            .zip(declarations)
                            .map(|((span, syntax), declaration)| Value::Declarator { span, declaration, syntax }),
                    );
                    values
                }
                _ => unreachable!(),
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_pratt_receipt_journal_restores_failed_rhs_and_every_budget() {
      for declaration_hooks in [true, false] {
        for source in ["a+b*", "a*b+", "f(a+b)*(c-d)+x[2*3]", "a+({int x=1+2; x;})*b"] {
            let mut builder = if declaration_hooks { SourcePrattBuilder::new(source) }
                else { SourcePrattBuilder::native(source) };
            let tree = super::super::build_hybrid(Rule::binary_expression, source,
                Default::default(), &mut builder).unwrap();
            let pair = tree.into_typed_pairs::<Rule>().next().unwrap();
            let reference = owned_reference(source);
            assert_eq!(pair.as_str(), reference.as_str());
            let span = pair.as_span();
            let recipe = builder.expressions.get(&(span.start(), span.end())).unwrap();
            let leaf = |child: grameo::compat::Pair<'_, Rule>| Expression::new(ExprKind::Ident(child.as_str().into()));
            let actual = recipe.lower(&pair, leaf).unwrap();
            let expected = construct_contextual_binary(&reference, leaf).unwrap();
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
        }
        let source = "int f(int a){return f(a+1)*(a-2)+3;} int broken( {";
        for max_steps in (0..512).chain(std::iter::once(usize::MAX)) {
            let mut builder = if declaration_hooks { SourcePrattBuilder::new(source) }
                else { SourcePrattBuilder::native(source) };
            assert!(super::super::build_hybrid(Rule::program, source,
                grameo::MatchOptions { max_steps, ..Default::default() }, &mut builder).is_err());
            assert!(builder.nodes.is_empty(), "budget={max_steps}");
            assert!(builder.roots.is_empty(), "budget={max_steps}");
            assert!(builder.side.is_empty(), "budget={max_steps}");
            assert!(builder.expressions.is_empty(), "budget={max_steps}");
            assert_eq!(builder.checkpoint(), 0, "budget={max_steps}");
        }
      }
    }

    #[test]
    fn native_source_pratt_reduction_graph_preserves_operator_schedules_and_handles() {
        let operators = ["||", "&&", "|", "^", "&", "==", "!=", "<", "<=", ">", ">=",
            "<<", ">>", "+", "-", "*", "/", "%"];
        for left in operators {
            for right in operators {
                let source = format!("++a {left} f(b+1) {right} c[2*3]");
                let mut builder = SourcePrattBuilder::native(&source);
                let tree = super::super::build_hybrid(Rule::binary_expression, &source,
                    Default::default(), &mut builder).unwrap();
                let pair = tree.into_typed_pairs::<Rule>().next().unwrap();
                let reference = owned_reference(&source);
                let leaf = |child: grameo::compat::Pair<'_, Rule>| Expression::new(
                    ExprKind::Ident(format!("{}:{}:{}", child.as_span().start(), child.as_span().end(), child.as_str())));
                let recipe = builder.expressions.get(&(0, pair.as_span().end())).unwrap();
                assert_eq!(format!("{:?}", recipe.lower(&pair, leaf).unwrap()),
                    format!("{:?}", construct_contextual_binary(&reference, leaf).unwrap()), "{source}");
                assert!(builder.into_declarations().is_empty());
            }
        }
    }

    fn owned_reference(source: &str) -> grameo::compat::Pair<'_, Rule> {
        grameo::engine::parse_program(&crate::native_grammar::Parser,
            Rule::binary_expression.name(), source, Default::default())
            .unwrap().into_typed_pairs().next().unwrap()
    }

    #[test]
    fn contextual_pratt_inline_tokens_preserve_spans_and_spill_boundary() {
        use grameo::pratt::{Token, TokenKind};
        for count in [0, 1, 3, 7, 8, 9, 17] {
            let mut tokens = ContextualTokens::new();
            for index in 0..count {
                tokens.push(Token {
                    kind: if index % 2 == 0 { TokenKind::Atom } else { TokenKind::Symbol("+") },
                    span: Span { start: index * 2, end: index * 2 + 1 },
                });
            }
            assert_eq!(tokens.spill.is_empty(), count <= 7);
            assert_eq!(tokens.as_slice().len(), count);
            for (index, token) in tokens.as_slice().iter().enumerate() {
                assert_eq!(token.span, Span { start: index * 2, end: index * 2 + 1 });
                assert!(matches!(token.kind, TokenKind::Atom) == (index % 2 == 0));
            }
        }
    }

    #[test]
    fn contextual_pratt_inline_and_spilled_regions_keep_complete_peg_ast() {
        for source in [
            "int f(int a,int b,int c,int d){return a+b*c-d;}",
            "int f(int a,int b,int c,int d,int e){return a+b*c-d+e;}",
            "typedef unsigned short word; int f(word *p,int a,int b,int c){return p[a]+sizeof(word)*b-c;}",
            "int side(int); int f(int a,int b,int c,int d,int e){return side(a+b)+side(c*d-e)+sizeof(int);}",
        ] {
            let expected = crate::parse_with_peg(source).unwrap();
            let actual = crate::parse_contextual_pratt(source).unwrap();
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
        }
    }

    #[test]
    fn contextual_pratt_single_infix_keeps_complete_peg_ast() {
        for operator in ["||", "&&", "|", "^", "&", "==", "!=", "<", "<=", ">", ">=",
            "<<", ">>", "+", "-", "*", "/", "%"] {
            let source = format!("int f(int a,int b){{return a /*left*/ {operator} /*right*/ b;}}");
            let expected = crate::parse_with_peg(&source).unwrap();
            let actual = crate::parse_contextual_pratt(&source).unwrap();
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
        }
        for source in [
            "typedef short word; int f(word *p,int i){return p[i++] + sizeof(word);}",
            "int side(int); int f(int a,int b){return side(a++) + side(++b);}",
            "int f(int a,int b){return (a+b) * (a-b);}",
            "int f(void){short a[8]; return a+2-a;}",
        ] {
            let expected = crate::parse_with_peg(source).unwrap();
            let actual = crate::parse_contextual_pratt(source).unwrap();
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
        }
    }

    #[test]
    fn contextual_pratt_inline_buffers_are_reentrant_and_keep_callback_order() {
        use super::super::{Backend, parse_captures, with_backend};
        for source in ["a+b", "a+b*c-d", "a+b*c-d+e"] {
            let pair = with_backend(Backend::Peg, ||
                parse_captures(Rule::binary_expression, source, Default::default()).unwrap().next().unwrap());
            let mut order = vec![];
            let result = construct_contextual_binary(&pair, |child| {
                if order.is_empty() {
                    let nested = with_backend(Backend::Peg, ||
                        parse_captures(Rule::binary_expression, "x+y", Default::default()).unwrap().next().unwrap());
                    let mut inner = vec![];
                    assert!(construct_contextual_binary(&nested, |value| {
                        inner.push(value.as_str().to_owned());
                        Expression::ident(value.as_str())
                    }).is_some());
                    assert_eq!(inner, ["x", "y"]);
                }
                order.push(child.as_str().to_owned());
                Expression::ident(child.as_str())
            });
            assert!(result.is_some());
            let expected: Vec<_> = source.split(['+', '*', '-']).map(str::to_owned).collect();
            assert_eq!(order, expected);
        }
    }

    #[test]
    fn contextual_pratt_prefix_lookahead_preserves_complete_operand_spans() {
        use super::super::{Backend, parse_captures, with_backend};
        for source in [
            "a",
            "a /*left*/ + /*right*/ b",
            "a+b*c",
            "a+b*c-d",
            "a+b*c-d+e",
            "a+b*c-d+e+f+g+h+i+j",
            "side(a++) + sizeof(int)",
        ] {
            let pair = with_backend(Backend::Peg, ||
                parse_captures(Rule::binary_expression, source, Default::default())
                    .unwrap().next().unwrap());
            let receipt = |child: &grameo::compat::Pair<'_, Rule>| {
                let span = child.as_span();
                (child.as_str().to_owned(), span.start(), span.end())
            };
            let expected: Vec<_> = pair.clone().into_inner()
                .filter(|child| child.as_rule() == Rule::unary_expression)
                .map(|child| receipt(&child)).collect();
            let mut actual = Vec::new();
            assert!(construct_contextual_binary(&pair, |child| {
                actual.push(receipt(&child));
                Expression::ident(child.as_str())
            }).is_some(), "{source}");
            assert_eq!(actual, expected, "{source}");
        }
        // A rejected boundary must not run a scope-sensitive operand callback.
        let pair = with_backend(Backend::Peg, ||
            parse_captures(Rule::program, "int value;", Default::default())
                .unwrap().next().unwrap());
        assert!(construct_contextual_binary(&pair, |_| {
            panic!("unvalidated operand callback");
        }).is_none());
    }

    #[test]
    fn numeric_hint_skips_comments_without_accepting_unterminated_trivia() {
        for source in ["1+2", "1 /*x*/ +2", "1//x\n+2", "1/*x*//*y*/\t+2"] {
            assert!(numeric_operator_hint(source), "{source}");
        }
        for source in ["1/*", "1//+2", "1/*+2*/", "1\\\n+2"] {
            assert!(!numeric_operator_hint(source), "{source}");
        }
    }

    fn operand(builder: &mut Owned<'_>, span: Span) {
        builder.begin(Rule::unary_expression as usize, span.start);
        builder.finish(Rule::int_literal as usize, span, builder.source);
        builder.finish(Rule::unary_expression as usize, span, builder.source);
    }

    #[test]
    fn repeated_rollback_restores_owned_children_and_folded_values() {
        for source in ["1+2", "1==2", "1<<40", "0xffffffffu+2"] {
            let mut builder = Owned::new(source);
            let (left_end, right_start) = if source.starts_with("0x") {
                (11, 12)
            } else if source.contains("==") || source.contains("<<") {
                (1, 3)
            } else {
                (1, 2)
            };
            operand(
                &mut builder,
                Span {
                    start: 0,
                    end: left_end,
                },
            );
            operand(
                &mut builder,
                Span {
                    start: right_start,
                    end: source.len(),
                },
            );
            let mark = builder.checkpoint();
            let snapshot = builder
                .arena
                .values()
                .iter()
                .map(|id| format!("{:?}", builder.arena.get(*id)))
                .collect::<Vec<_>>();
            for _ in 0..3 {
                builder.reduce(
                    0,
                    grameo::pratt::Fixity::InfixLeft,
                    Span {
                        start: left_end,
                        end: right_start,
                    },
                );
                assert_eq!(builder.arena.values().len(), 1);
                builder.rollback(mark);
                assert_eq!(
                    snapshot,
                    builder
                        .arena
                        .values()
                        .iter()
                        .map(|id| format!("{:?}", builder.arena.get(*id)))
                        .collect::<Vec<_>>()
                );
            }
        }
    }

    #[test]
    fn prefix_composition_restores_literal_nodes_and_owner_frames() {
        for source in ["-2", "+2", "!2", "~2"] {
            let mut builder = Owned::new(source);
            builder.begin(Rule::binary_expression as usize, 0);
            builder.begin(Rule::unary_expression as usize, 0);
            builder.finish(Rule::prefix_op as usize, Span { start: 0, end: 1 }, source);
            operand(&mut builder, Span { start: 1, end: 2 });
            let mark = builder.checkpoint();
            let snapshot = builder
                .arena
                .values()
                .iter()
                .map(|id| format!("{:?}", builder.arena.get(*id)))
                .collect::<Vec<_>>();
            for _ in 0..3 {
                builder.finish(
                    Rule::unary_expression as usize,
                    Span { start: 0, end: 2 },
                    source,
                );
                assert!(
                    builder
                        .accepts_pratt(Rule::binary_expression as usize, Span { start: 0, end: 2 })
                );
                builder.rollback(mark);
                assert_eq!(
                    snapshot,
                    builder
                        .arena
                        .values()
                        .iter()
                        .map(|id| format!("{:?}", builder.arena.get(*id)))
                        .collect::<Vec<_>>()
                );
            }
        }
    }

    #[test]
    fn mixed_region_composition_is_private_and_reverses_exactly() {
        let source = "x+2";
        let mut builder = MixedBuilder::default();
        builder.begin(Rule::binary_expression as usize, 0);
        let before_operator = builder.checkpoint();
        builder.finish(Rule::binary_op as usize, Span { start: 1, end: 2 }, source);
        builder.rollback(before_operator);
        assert_eq!(builder.arena.values().len(), 1);
        assert!(builder.observes_rule(Rule::binary_expression as usize));
        builder.finish(Rule::binary_op as usize, Span { start: 1, end: 2 }, source);
        let mark = builder.checkpoint();
        for _ in 0..3 {
            builder.finish(
                Rule::binary_expression as usize,
                Span { start: 0, end: 3 },
                source,
            );
            assert!(builder.expressions.is_empty());
            assert_eq!(builder.arena.values().len(), 1);
            builder.rollback(mark);
            assert_eq!(builder.arena.values().len(), 2);
            assert!(matches!(
                builder.arena.get(builder.arena.values()[1]),
                MixedValue::Operator(BinOp::Add)
            ));
        }
        builder.rollback(0);
        assert_eq!(builder.checkpoint(), 0);
        assert!(builder.arena.values().is_empty());
        assert!(builder.expressions.is_empty());
    }

    #[test]
    fn singleton_mixed_region_owner_frame_restores_without_recipe() {
        let mut builder = MixedBuilder::default();
        builder.begin(Rule::binary_expression as usize, 0);
        let mark = builder.checkpoint();
        for _ in 0..3 {
            builder.finish(
                Rule::binary_expression as usize,
                Span { start: 0, end: 1 },
                "x",
            );
            assert!(builder.arena.values().is_empty());
            assert!(!builder.observes_rule(Rule::unary_expression as usize));
            builder.rollback(mark);
            assert_eq!(builder.arena.values().len(), 1);
            assert!(matches!(
                builder.arena.get(builder.arena.values()[0]),
                MixedValue::Frame(0)
            ));
        }
        builder.finish(
            Rule::binary_expression as usize,
            Span { start: 0, end: 1 },
            "x",
        );
        builder.committed("x", 1);
        assert!(builder.expressions.is_empty());
    }

    #[test]
    fn declaration_bindings_rollback_after_reductions_and_late_failure() {
        use declarations::{
            BINDINGS, DeclarationContext, DeclarationPolicy, DeclarationTarget, ScalarPolicy,
        };
        use grameo::ast::builder::AstBuilder;
        for source in [
            "int x=2; int y=;",
            "int x=-2+3*4; int y=;",
            "int x=1==2; int y=;",
            "int x=1<<40; int y=;",
            "int x=L'\\n'; int y=;",
        ] {
            for max_steps in [20, usize::MAX] {
                let mut context = DeclarationContext::new(source);
                let end = source.find(';').unwrap();
                context.declarations.insert(
                    (0, end + 1),
                    DeclarationPolicy {
                        kind: vybe_ast::VarDeclKind::Var,
                        statement_span: Default::default(),
                    },
                );
                context.declarators.insert(
                    (4, end),
                    ScalarPolicy {
                        emitted_name: "renamed_x".into(),
                        type_hint: Some("int".into()),
                        zero_if_missing: true,
                        integer_wrap: None,
                    },
                );
                let mut builder = AstBuilder::<DeclarationTarget>::with_context(BINDINGS, context);
                let initial = builder.checkpoint();
                assert!(
                    super::super::build_hybrid(
                        Rule::program,
                        source,
                        grameo::MatchOptions {
                            max_steps,
                            ..Default::default()
                        },
                        &mut builder
                    )
                    .is_err()
                );
                assert_eq!(builder.checkpoint(), initial);
                assert!(builder.into_value().is_err());
            }
        }
    }

    #[test]
    fn declaration_binary_inverse_restores_folded_and_moved_children() {
        use declarations::{DeclarationContext, DeclarationTarget, Value as DeclValue};
        use grameo::ast::builder::{Argument, Target};
        for (op, left, right) in [
            (BinOp::Eq, "1", "2"),
            (BinOp::Shl, "1", "40"),
            (BinOp::Add, "1u", "2"),
        ] {
            let children = vec![
                DeclValue::Scalar {
                    span: Span { start: 0, end: 2 },
                    expression: integer(left),
                },
                DeclValue::Operator {
                    span: Span { start: 2, end: 3 },
                    op,
                },
                DeclValue::Scalar {
                    span: Span { start: 3, end: 5 },
                    expression: integer(right),
                },
            ];
            let before = format!("{children:?}");
            let address = match &children[0] {
                DeclValue::Scalar {
                    expression:
                        Expression {
                            kind: ExprKind::Cast { expr, .. },
                            ..
                        },
                    ..
                } => Some(&**expr as *const Expression),
                _ => None,
            };
            let (value, undo) = DeclarationTarget::construct(
                &DeclarationContext::default(),
                14,
                Span { start: 0, end: 5 },
                vec![Argument::Many(children)],
            );
            let children = DeclarationTarget::split(14, value, undo);
            assert_eq!(format!("{children:?}"), before);
            if let Some(address) = address {
                let DeclValue::Scalar {
                    expression:
                        Expression {
                            kind: ExprKind::Cast { expr, .. },
                            ..
                        },
                    ..
                } = &children[0]
                else {
                    unreachable!()
                };
                assert_eq!(address, &**expr as *const Expression);
            }
        }
    }

    #[test]
    fn declaration_inverse_returns_moved_initializer_and_original_name() {
        use declarations::{
            DeclarationContext, DeclarationTarget, ScalarPolicy, Undo as DeclUndo, UnsignedWidth,
            Value as DeclValue,
        };
        use grameo::ast::builder::{Argument, Target};
        for integer_wrap in [None, Some(UnsignedWidth::W32)] {
            let span = Span { start: 4, end: 8 };
            let mut context = DeclarationContext::default();
            context.declarators.insert(
                (4, 8),
                ScalarPolicy {
                    emitted_name: "renamed_x".into(),
                    type_hint: Some("int".into()),
                    zero_if_missing: true,
                    integer_wrap,
                },
            );
            let value = integer("1u");
            let ExprKind::Cast { expr, .. } = &value.kind else {
                unreachable!()
            };
            let address = &**expr as *const Expression;
            let (value, undo) = DeclarationTarget::construct(
                &context,
                5,
                span,
                vec![Argument::Many(vec![
                    DeclValue::PlainName {
                        span: Span { start: 4, end: 5 },
                        name_span: Span { start: 4, end: 5 },
                        text: "x".into(),
                    },
                    DeclValue::Initializer {
                        span: Span { start: 6, end: 8 },
                        literal_span: Span { start: 6, end: 8 },
                        expression: value,
                    },
                ])],
            );
            assert!(matches!(undo, DeclUndo::Declarator { .. }));
            let mut children = DeclarationTarget::split(5, value, undo);
            let DeclValue::Initializer { expression, .. } = children.pop().unwrap() else {
                unreachable!()
            };
            let ExprKind::Cast { expr, .. } = &expression.kind else {
                unreachable!()
            };
            assert_eq!(address, &**expr as *const Expression);
            assert!(
                matches!(children.pop().unwrap(), DeclValue::PlainName { text, .. } if text == "x")
            );
        }
    }

    #[test]
    fn mixed_matching_failure_rolls_back_finished_regions_and_atoms() {
        for (source, max_steps) in [
            (
                "int f(void){return x+2;} int g(void){return y+;}",
                usize::MAX,
            ),
            ("int f(void){return x+2;}", 20),
        ] {
            let mut builder = MixedBuilder::default();
            assert!(
                super::super::build_hybrid(
                    Rule::program,
                    source,
                    grameo::MatchOptions {
                        max_steps,
                        ..Default::default()
                    },
                    &mut builder
                )
                .is_err()
            );
            assert_eq!(builder.checkpoint(), 0);
            assert!(builder.arena.values().is_empty());
            assert!(builder.expressions.is_empty());
        }
    }

    #[test]
    fn combined_declaration_and_binary_journals_restore_every_failed_budget() {
        for source in [
            "int x=2; int f(void){int y=3; return x+y;} int bad=;",
            "int x[2]={1,2}; int f(void){int y=3; return x[0]+y;} int bad=;",
            "int x=({int y=3; y;}); int bad=;",
            "int a=1,*p=0,b=2; int f(void){int c=3,d[2]={1,2},e=4; return c+e;} int bad=;",
        ] {
            for max_steps in [1, 20, 100, 1000, usize::MAX] {
                let mut builder = ContextualBuilder::new(source);
                assert!(
                    super::super::build_hybrid(
                        Rule::program,
                        source,
                        grameo::MatchOptions {
                            max_steps,
                            ..Default::default()
                        },
                        &mut builder
                    )
                    .is_err()
                );
                assert_eq!(builder.checkpoint(), 0);
                assert_eq!(builder.depth, 0);
                assert_eq!(builder.arrays, 0);
                assert!(builder.mixed.arena.values().is_empty());
                assert!(builder.declarations.is_empty());
                assert!(builder.mixed.expressions.is_empty());
            }
        }
    }

    #[test]
    fn plain_scalar_receipt_inverse_preserves_each_declarator_syntax() {
        use declarations::{DeclarationContext, DeclarationTarget, Value as DeclValue};
        use grameo::ast::builder::{Argument, Target};
        let source = "int a=1,b=2;";
        let mut children = vec![DeclValue::Specs {
            span: Span { start: 0, end: 3 }, children: vec![],
        }];
        for (name, start, value) in [("a", 4, 1), ("b", 8, 2)] {
            children.push(DeclValue::Declarator {
                span: Span { start, end: start + 3 },
                syntax: PlainDeclaratorSyntax {
                    declarator: Span { start, end: start + 1 },
                    initializer: Some(Span { start: start + 2, end: start + 3 }),
                },
                declaration: vybe_ast::VarDeclarator {
                    pattern: vybe_ast::BindingPattern::Ident(name.into()),
                    type_hint: None, init: Some(int(value)), array_bounds: None,
                    with_events: false,
                },
            });
        }
        let expected = format!("{children:?}");
        for _ in 0..3 {
            let (value, undo) = DeclarationTarget::construct(
                &DeclarationContext::deferred(source), 6,
                Span { start: 0, end: source.len() }, vec![Argument::Many(children)],
            );
            assert!(matches!(&value, DeclValue::Declaration { .. }));
            children = DeclarationTarget::split(6, value, undo);
            assert_eq!(format!("{children:?}"), expected);
        }
    }

    #[test]
    fn plain_scalar_receipt_name_tail_requires_exact_grammar_trivia() {
        use declarations::{DeclarationContext, DeclarationTarget, Value as DeclValue};
        use grameo::ast::builder::{Argument, Target};
        for (tail, accepted) in [
            ("", true), (" \t\r\n", true), (" /* gap */ ", true),
            (" // gap\r\n", true), ("\\\n /* gap */", true),
            ("\\\r\n", true), ("/*", false), ("\\\r", false),
            ("[]", false), ("()", false), ("*", false), (";", false),
            (" \u{b}", false), ("\"/* gap */\"", false),
        ] {
            let source = format!("name{tail}");
            let (value, undo) = DeclarationTarget::construct(
                &DeclarationContext::deferred(&source), 3,
                Span { start: 0, end: source.len() },
                vec![Argument::Many(vec![DeclValue::Name {
                    span: Span { start: 0, end: 4 }, text: "name".into(),
                }])],
            );
            assert_eq!(matches!(&value, DeclValue::PlainName { .. }), accepted, "{tail:?}");
            let children = DeclarationTarget::split(3, value, undo);
            assert!(matches!(children.as_slice(), [DeclValue::Name { span, text }]
                if *span == Span { start: 0, end: 4 } && text == "name"));
        }
    }

    #[test]
    fn unsupported_operand_cannot_consume_a_ready_sibling() {
        let mut builder = Owned::new("1+x");
        operand(&mut builder, Span { start: 0, end: 1 });
        builder.begin(Rule::unary_expression as usize, 2);
        builder.finish(
            Rule::unary_expression as usize,
            Span { start: 2, end: 3 },
            "1+x",
        );
        let mark = builder.checkpoint();
        builder.reduce(
            0,
            grameo::pratt::Fixity::InfixLeft,
            Span { start: 1, end: 2 },
        );
        assert_eq!(builder.checkpoint(), mark);
        assert!(
            !builder.accepts_pratt(Rule::binary_expression as usize, Span { start: 0, end: 3 })
        );
    }

    #[test]
    fn integer_suffix_widths_values_and_default_spans_are_exact() {
        for (raw, value, ty) in [
            ("2147483648", 2147483648, None),
            ("4294967295", 4294967295, None),
            ("9223372036854775808", i64::MIN, None),
            ("18446744073709551615ULL", -1, Some("uint64")),
            ("18446744073709551616ULL", 0, Some("uint64")),
            ("0x80000000", 2147483648, Some("uint32")),
            ("0xffffffff", 4294967295, Some("uint32")),
            ("037777777777", 4294967295, Some("uint32")),
            ("01777777777777777777777", -1, None),
            ("08", 0, None),
            ("0xffffffffffffffff", -1, None),
            ("0x10000000000000000", 0, None),
            (
                "0b1111111111111111111111111111111111111111111111111111111111111111",
                -1,
                None,
            ),
            ("1u", 1, Some("uint32")),
            ("1U", 1, Some("uint32")),
            ("1l", 1, Some("int64")),
            ("1L", 1, Some("int64")),
            ("1ll", 1, Some("int64")),
            ("1LL", 1, Some("int64")),
            ("1uL", 1, Some("uint64")),
            ("1Lu", 1, Some("uint64")),
            ("1UlL", 1, Some("uint64")),
            ("0xffffffffffffffffULL", -1, Some("uint64")),
        ] {
            let literal = int(value);
            let expected = match ty {
                Some(type_name) => Expression::new(ExprKind::Cast {
                    expr: Box::new(literal),
                    type_name: type_name.into(),
                }),
                None => literal,
            };
            assert_eq!(
                format!("{:?}", integer(raw)),
                format!("{expected:?}"),
                "{raw}"
            );
        }
    }

    #[test]
    fn pratt_acceptance_requires_the_matching_owner_frame() {
        let mut builder = Owned::new("1");
        operand(&mut builder, Span { start: 0, end: 1 });
        assert!(
            !builder.accepts_pratt(Rule::binary_expression as usize, Span { start: 0, end: 1 })
        );
        let mut builder = Owned::new("1");
        builder.begin(Rule::binary_expression as usize, 0);
        operand(&mut builder, Span { start: 0, end: 1 });
        assert!(builder.accepts_pratt(Rule::binary_expression as usize, Span { start: 0, end: 1 }));
        assert!(!builder.accepts_pratt(Rule::unary_expression as usize, Span { start: 0, end: 1 }));
    }

    #[test]
    fn completed_island_remains_private_and_first_handoff_moves_children() {
        let source = "1u+2LL";
        let mut builder = Owned::new(source);
        let initial = builder.checkpoint();
        for attempt in 0..2 {
            builder.begin(Rule::binary_expression as usize, 0);
            operand(&mut builder, Span { start: 0, end: 2 });
            operand(
                &mut builder,
                Span {
                    start: 3,
                    end: source.len(),
                },
            );
            builder.reduce(
                0,
                grameo::pratt::Fixity::InfixLeft,
                Span { start: 2, end: 3 },
            );
            builder.finish(
                Rule::binary_expression as usize,
                Span {
                    start: 0,
                    end: source.len(),
                },
                source,
            );
            assert!(
                builder.expressions.is_empty(),
                "finish published a speculative island"
            );
            if attempt == 0 {
                builder.rollback(initial);
                assert!(builder.arena.values().is_empty());
            }
        }
        let id = *builder.arena.values().last().unwrap();
        let Value::Operand {
            expression: Some(expression),
            ..
        } = builder.arena.get(id)
        else {
            panic!("missing completed expression");
        };
        let child_addresses = |expression: &Expression| {
            let ExprKind::Binary { left, right, .. } = &expression.kind else {
                panic!("expected binary tree");
            };
            (
                &**left as *const Expression as usize,
                &**right as *const Expression as usize,
            )
        };
        let addresses = child_addresses(expression);
        let expected = format!("{expression:?}");
        builder.committed(source, source.len());
        let mut cached = builder.expressions.remove(&(0, source.len())).unwrap();
        let mut first = cached.take(source);
        assert_eq!(
            addresses,
            child_addresses(&first),
            "first handoff cloned child allocations"
        );
        assert!(cached.first.is_none());
        // A consuming normalization of the first tree must not alter replay.
        first.kind = ExprKind::Lit(Literal::Int(99));
        assert_eq!(format!("{first:?}"), format!("{:?}", int(99)));
        assert_eq!(expected, format!("{:?}", cached.take(source)));
        assert_eq!(expected, format!("{:?}", cached.take(source)));
    }

    #[test]
    fn later_function_failure_rolls_back_earlier_finished_islands() {
        let source = "int f(void) { return 1u+2LL; } int g(void) { return 3+; }";
        let mut builder = Owned::new(source);
        let result = super::super::build_hybrid(
            Rule::program,
            source,
            grameo::MatchOptions::default(),
            &mut builder,
        );
        assert!(result.is_err());
        assert_eq!(builder.checkpoint(), 0);
        assert!(builder.arena.values().is_empty());
        assert!(
            builder.expressions.is_empty(),
            "earlier function committed before module success"
        );
    }

    #[test]
    fn generated_pratt_constructs_owned_ast_and_replays_exactly() {
        for source in [
            "0",
            "0xffffffffu",
            "1+2*3",
            "1==2",
            "1<<40",
            "0xffffffffu+2",
            "1<2<<3",
            "1|2||3",
            "1 /*x*/ + 2",
            "1 //x\n + 2",
            "1+2*3==7!=0",
            "0x80000000+1LL",
            "0xffffffffffffffffULL>>1",
            "-2",
            "+2",
            "!2",
            "~2",
            "-~!+2",
            "1+-2",
            "-1+~2*3",
            "1+2.0",
            "0x1.8p+2 + .5f",
            "'a'+'\\n'",
            "L'\\u00ff'+u8'\\x41'",
            "true==false",
        ] {
            let mut builder = Owned::new(source);
            let tree = super::super::build_hybrid(
                Rule::binary_expression,
                source,
                grameo::MatchOptions::default(),
                &mut builder,
            )
            .unwrap();
            assert_eq!(tree.consumed(), source.len(), "{source}");
            let mut cached = builder.expressions.remove(&(0, source.len())).unwrap();
            let first = cached.take(source);
            assert_eq!(
                format!("{first:?}"),
                format!("{:?}", cached.take(source)),
                "{source}"
            );
            assert_eq!(
                format!("{first:?}"),
                format!("{:?}", cached.take(source)),
                "{source}"
            );
            let root = tree.into_typed_pairs::<Rule>().next().unwrap();
            assert_eq!(
                root.into_inner().count(),
                0,
                "owned island retained descendant captures: {source}"
            );
        }
    }

    #[test]
    fn rich_unary_grammar_recaptures_transactionally() {
        for source in [
            "1+x",
            "1+2*3+x",
            "1==2||x",
            "1<<40|x",
            "0xffffffffu+2+x",
            "1+(2+3)",
            "1+(int)x",
            "1+f(2)",
            "1+p->x",
            "1+sizeof(int)",
            "1+sizeof(2==3)",
            "1+_Alignof(int)",
            "1+va_arg(ap,int)",
            "1+_Generic(x,int:2,default:3)",
            "1++",
            "1(2)",
            "1(2+3)",
            "1(2(3))",
            "1(2+sizeof(3))",
            "1[0]",
            "1[2+3]",
            "-x",
            "~x",
            "!p",
            "++2",
            "--2",
            "*2",
            "&2",
            "1+*2",
            "1+&(2+3)",
            "1\\\n+2",
        ] {
            let mut builder = Owned::new(source);
            let options = grameo::MatchOptions::default();
            let candidate =
                super::super::build_hybrid(Rule::binary_expression, source, options, &mut builder)
                    .unwrap();
            let reference = grameo::engine::parse_program(
                &crate::native_grammar::Parser,
                Rule::binary_expression.name(),
                source,
                options,
            )
            .unwrap();
            let flatten = |tree: grameo::ParseTree<'_, '_>| {
                tree.into_typed_pairs::<Rule>()
                    .flatten()
                    .map(|pair| (pair.as_rule(), pair.as_span().start(), pair.as_span().end()))
                    .collect::<Vec<_>>()
            };
            assert_eq!(flatten(reference), flatten(candidate), "{source}");
            assert!(
                builder.expressions.is_empty(),
                "unsupported island published AST: {source}"
            );
        }
    }

    #[test]
    fn contextual_postfix_keeps_independent_integer_islands() {
        for source in ["f(1+2)", "x[1+2]"] {
            let mut builder = Owned::new(source);
            let tree = super::super::build_hybrid(
                Rule::binary_expression,
                source,
                grameo::MatchOptions::default(),
                &mut builder,
            )
            .unwrap();
            assert_eq!(tree.consumed(), source.len(), "{source}");
            assert_eq!(builder.expressions.len(), 1, "{source}");
            let mut cached = builder.expressions.remove(&(2, 5)).unwrap();
            let expected = combine(BinOp::Add, int(1), int(2));
            assert_eq!(
                format!("{:?}", cached.take(source)),
                format!("{expected:?}")
            );
            assert_eq!(
                format!("{:?}", cached.take(source)),
                format!("{expected:?}")
            );
        }
    }

    #[test]
    fn postfix_exclusion_rolls_back_frames_and_contextual_boundaries() {
        let source = "1(2)";
        let mut builder = Owned::new(source);
        builder.begin(Rule::binary_expression as usize, 0);
        builder.begin(Rule::unary_expression as usize, 0);
        builder.finish(
            Rule::int_literal as usize,
            Span { start: 0, end: 1 },
            source,
        );
        let initial = builder.checkpoint();
        let snapshot = builder
            .arena
            .values()
            .iter()
            .map(|id| format!("{:?}", builder.arena.get(*id)))
            .collect::<Vec<_>>();
        for _ in 0..3 {
            builder.begin(Rule::call_suffix as usize, 1);
            let inside = builder.checkpoint();
            assert_eq!(builder.blocked, 1);
            assert!(!builder.supports_pratt_at(Rule::binary_expression as usize, 0, "2+3"));
            assert!(!builder.observes_rule(Rule::call_suffix as usize));
            builder.begin(Rule::sizeof_expression as usize, 2);
            assert_eq!(builder.blocked, 2);
            builder.rollback(inside);
            assert_eq!(builder.blocked, 1);
            builder.finish(
                Rule::call_suffix as usize,
                Span { start: 1, end: 4 },
                source,
            );
            assert_eq!(builder.blocked, 0);
            builder.finish(
                Rule::unary_expression as usize,
                Span { start: 0, end: 4 },
                source,
            );
            assert!(
                !builder.accepts_pratt(Rule::binary_expression as usize, Span { start: 0, end: 4 })
            );
            builder.rollback(inside);
            assert_eq!(builder.blocked, 1);
            builder.rollback(initial);
            assert_eq!(builder.blocked, 0);
            assert_eq!(builder.checkpoint(), initial);
            assert_eq!(
                snapshot,
                builder
                    .arena
                    .values()
                    .iter()
                    .map(|id| format!("{:?}", builder.arena.get(*id)))
                    .collect::<Vec<_>>()
            );
        }
    }

    #[test]
    fn failed_and_budgeted_matches_publish_no_owned_nodes() {
        for (source, max_steps) in [
            ("1+2", 1),
            ("1+2", 20),
            ("1(2+3)", 20),
            ("1[2+3]", 20),
            ("-2+3", 20),
            ("1+*2", 20),
            ("+", 1_000_000),
        ] {
            let mut builder = Owned::new(source);
            let result = super::super::build_hybrid(
                Rule::binary_expression,
                source,
                grameo::MatchOptions {
                    max_steps,
                    ..Default::default()
                },
                &mut builder,
            );
            assert!(result.is_err(), "{source} budget={max_steps}");
            assert_eq!(builder.checkpoint(), 0);
            assert!(builder.arena.values().is_empty());
            assert!(builder.expressions.is_empty());
        }
    }

    #[test]
    fn contextual_exclusions_restore_every_interleaved_boundary() {
        let mut builder = Owned::new("1");
        let outer = builder.checkpoint();
        builder.begin(Rule::sizeof_expression as usize, 0);
        let first = builder.checkpoint();
        builder.begin(Rule::type_name as usize, 0);
        let second = builder.checkpoint();
        assert_eq!(builder.blocked, 2);
        builder.finish(Rule::type_name as usize, Span { start: 0, end: 1 }, "1");
        builder.rollback(second);
        assert_eq!(builder.blocked, 2);
        builder.rollback(first);
        assert_eq!(builder.blocked, 1);
        builder.finish(
            Rule::sizeof_expression as usize,
            Span { start: 0, end: 1 },
            "1",
        );
        operand(&mut builder, Span { start: 0, end: 1 });
        builder.rollback(first);
        assert_eq!(builder.blocked, 1);
        assert!(builder.arena.values().is_empty());
        builder.rollback(outer);
        assert_eq!(builder.blocked, 0);
        assert_eq!(builder.checkpoint(), 0);
    }
}
