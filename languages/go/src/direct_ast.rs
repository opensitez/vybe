//! Go expressions constructed from the grammar's own hooks.
//!
//! [`ExpressionBuilder`] journals the begin/finish hooks of the rules it
//! observes and the reductions of the Go Pratt islands in one append-only
//! log: a checkpoint is the log length and a rollback truncates it, so a
//! failed alternative leaves no trace. When an expression rule finishes, its
//! own events are replayed by a stack machine and replaced by one event that
//! holds the result, so every event is replayed once.
//!
//! A region is owned when every node inside it is one this module builds:
//! identifiers, literals, unary operators, binary operators (grouped by the
//! Pratt reductions, which carry Go's five precedence levels) and
//! parenthesized expressions. An owned expression's child captures are
//! pruned: the walker receives the expression instead. Any other construct
//! makes a region incomplete; the walker lowers it from its captures, and the
//! complete expressions nested inside it remain owned.
//!
//! Owned expressions are keyed by their source, rule and byte span.

use crate::Rule;
use grameo::Span;
use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use vybe_ast::{BinOp, ExprKind, Expression, Literal, UnaryOp};

type RegionKey = (usize, usize, usize);
type SourceKey = (usize, usize);

thread_local! {
    static OWNED: RefCell<HashMap<SourceKey, HashMap<RegionKey, Expression>>> =
        RefCell::new(HashMap::new());
    static PROFILE: Cell<ConstructionProfile> = const { Cell::new(ConstructionProfile::ZERO) };
}

/// Expression construction counts for the Go parses on this thread.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct ConstructionProfile {
    /// Expression roots the builder constructed completely.
    pub owned_regions: usize,
    /// Expression roots the walker received from the builder.
    pub handoffs: usize,
    /// Expression roots the walker lowered from captures.
    pub walked: usize,
}

impl ConstructionProfile {
    const ZERO: Self = Self { owned_regions: 0, handoffs: 0, walked: 0 };
}

fn record(change: impl FnOnce(&mut ConstructionProfile)) {
    PROFILE.with(|cell| {
        let mut profile = cell.get();
        change(&mut profile);
        cell.set(profile);
    });
}

/// Run `action` and return the construction counts it accumulated.
pub fn profile_construction<T>(action: impl FnOnce() -> T) -> (T, ConstructionProfile) {
    let before = PROFILE.with(|cell| cell.replace(ConstructionProfile::ZERO));
    let value = action();
    let counted = PROFILE.with(|cell| cell.get());
    PROFILE.with(|cell| {
        cell.set(ConstructionProfile {
            owned_regions: before.owned_regions + counted.owned_regions,
            handoffs: before.handoffs + counted.handoffs,
            walked: before.walked + counted.walked,
        })
    });
    (value, counted)
}

/// Discard every owned expression on this thread. Lookups then miss and the
/// walker lowers those regions from their captures.
pub(crate) fn clear_owned() {
    OWNED.with(|owned| owned.borrow_mut().clear());
}

fn source_key(source: &str) -> SourceKey {
    (source.as_ptr() as usize, source.len())
}

/// The public parse route selects the expression builder only when it costs
/// no more retired instructions than the reference builder; it currently
/// does not, so parsing uses the reference builder and the walker lowers
/// every expression.
const EXPRESSION_BUILDER_SELECTED: bool = false;

fn is_expression_root(rule: usize) -> bool {
    rule == Rule::expression as usize
        || rule == Rule::if_expression as usize
        || rule == Rule::range_expression as usize
}

/// Parse `source` from `rule` without semantic observations.
pub(crate) fn parse_reference(
    rule: Rule,
    source: &str,
) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
    crate::native_grammar::Parser
        .build_captures(rule, source, &mut ReferenceBuilder)
        .map(|tree| tree.into_typed_pairs())
}

/// Parse `source` from `rule` and record the expressions the grammar hooks
/// constructed. Their captures keep only the expression rule itself, which
/// the walker exchanges for the owned value.
pub(crate) fn parse_owned(
    rule: Rule,
    source: &str,
) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
    if !EXPRESSION_BUILDER_SELECTED {
        return parse_reference(rule, source);
    }
    let key = source_key(source);
    OWNED.with(|owned| owned.borrow_mut().remove(&key));
    let mut builder = ExpressionBuilder::new(source);
    let tree = match crate::native_grammar::Parser.build_captures(rule, source, &mut builder) {
        Ok(tree) => tree,
        // The reference grammar reports the syntax error: its expected-token
        // set lists the binary operators that the Pratt island tries silently.
        Err(_) => return parse_reference(rule, source),
    };
    let regions = builder.into_regions();
    record(|profile| profile.owned_regions += regions.len());
    if !regions.is_empty() {
        OWNED.with(|owned| owned.borrow_mut().insert(key, regions));
    }
    Ok(tree.into_typed_pairs())
}

/// The owned expression for this expression capture. An expression capture
/// without children was pruned for an owned value, so it must have one.
pub(crate) fn owned_expression(
    pair: &grameo::compat::Pair<'_, Rule>,
) -> Result<Option<Expression>, String> {
    let rule = pair.as_rule() as usize;
    if !EXPRESSION_BUILDER_SELECTED || !is_expression_root(rule) {
        return Ok(None);
    }
    let span = pair.as_span();
    let owned = OWNED.with(|owned| {
        owned
            .borrow()
            .get(&source_key(span.get_input()))?
            .get(&(rule, span.start(), span.end()))
            .cloned()
    });
    record(|profile| {
        if owned.is_some() {
            profile.handoffs += 1;
        } else {
            profile.walked += 1;
        }
    });
    if owned.is_none() && pair.clone().into_inner().next().is_none() {
        return Err(format!("Go expression `{}` has neither captures nor an owned value", pair.as_str()));
    }
    Ok(owned)
}

// ── Builder ─────────────────────────────────────────────────────────────────

#[derive(Clone, Copy)]
enum Event {
    Begin,
    Finish(usize, Span),
    Reduce(Span),
    /// A complete expression region; its value is in the referenced slot.
    Root(usize),
    /// A complete expression region inside an incomplete one: the walker
    /// receives it, and no enclosing construct consumes it.
    Released(usize),
    /// An incomplete expression region.
    Incomplete,
}

/// The grammar's silent newline trivia (`nls`), when the grammar has one. No
/// builder observes it, so omitting its certified empty matches is inert.
fn newline_trivia_rule() -> Option<usize> {
    use grameo::program::Program;
    static RULE: std::sync::OnceLock<Option<usize>> = std::sync::OnceLock::new();
    *RULE.get_or_init(|| crate::native_grammar::Parser.rule_id("nls"))
}

/// The grammar without semantic observations or Pratt islands.
struct ReferenceBuilder;

impl grameo::builder::Builder for ReferenceBuilder {
    fn supports_choice_heads(&self) -> bool { true }
    fn supports_unobserved_failure_prefixes(&self) -> bool { true }
    fn observes_rule(&self, _: usize) -> bool { false }
    fn opaque_rule(&self, _: usize) -> bool { true }
    fn empty_success_is_inert(&self, rule: usize) -> bool { Some(rule) == newline_trivia_rule() }
    fn checkpoint(&self) -> usize { 0 }
    fn rollback(&mut self, _: usize) {}
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
}

struct Slot {
    /// Index of the event that refers to this slot.
    event: usize,
    key: RegionKey,
    expr: Option<Expression>,
}

struct ExpressionBuilder<'s> {
    source: &'s str,
    log: Vec<Event>,
    slots: Vec<Slot>,
}

impl<'s> ExpressionBuilder<'s> {
    fn new(source: &'s str) -> Self {
        Self { source, log: Vec::new(), slots: Vec::new() }
    }

    fn slot(&mut self, key: RegionKey, expr: Expression) -> usize {
        self.slots.push(Slot { event: self.log.len(), key, expr: Some(expr) });
        self.slots.len() - 1
    }

    /// Replace a finished expression rule's events by its result.
    fn collapse(&mut self, rule: usize, span: Span) {
        let mut depth = 0usize;
        let mut begin = self.log.len();
        while begin > 0 {
            begin -= 1;
            match self.log[begin] {
                Event::Finish(..) => depth += 1,
                Event::Begin if depth == 0 => break,
                Event::Begin => depth -= 1,
                _ => {}
            }
        }
        let mut replay = Replay::new(self.source);
        for index in begin + 1..self.log.len() {
            let event = self.log[index];
            replay.event(event, &mut self.slots);
        }
        let children = replay.remaining();
        let value = replay.root(rule, span, children);
        let released = std::mem::take(&mut replay.released);
        self.log.truncate(begin);
        while self.slots.last().is_some_and(|slot| slot.event >= begin) {
            self.slots.pop();
        }
        for (key, expr) in released {
            let id = self.slot(key, expr);
            self.log.push(Event::Released(id));
        }
        match value {
            Value::Expr { expr, .. } => {
                let id = self.slot((rule, span.start, span.end), expr);
                self.log.push(Event::Root(id));
            }
            Value::Incomplete => self.log.push(Event::Incomplete),
        }
    }

    fn into_regions(mut self) -> HashMap<RegionKey, Expression> {
        let mut regions = HashMap::new();
        for event in &self.log {
            if let Event::Root(id) | Event::Released(id) = *event {
                let slot = &mut self.slots[id];
                if let Some(expr) = slot.expr.take() {
                    regions.insert(slot.key, expr);
                }
            }
        }
        regions
    }
}

/// Rules whose hooks the replay reads. Each is reached only where its
/// alternatives already matched or are about to, so observing them leaves the
/// certified choice-head and failed-prefix shortcuts of every other rule
/// intact. Unobserved constructs are recognized from the observed spans.
static OBSERVED: [bool; RULE_TABLE] = observed_rules();

/// An upper bound on the grammar's rule count; `observed_rules` fails to
/// compile when a rule lies beyond it.
const RULE_TABLE: usize = 256;

const fn observed_rules() -> [bool; RULE_TABLE] {
    use Rule as R;
    let rules = [
        R::expression, R::if_expression, R::range_expression,
        R::unary_expression, R::if_unary_expression, R::range_unary_expression,
        R::primary, R::if_primary, R::range_primary,
        R::operand, R::if_operand, R::range_operand,
    ];
    let mut table = [false; RULE_TABLE];
    let mut index = 0;
    while index < rules.len() {
        table[rules[index] as usize] = true;
        index += 1;
    }
    table
}

fn observed(rule: usize) -> bool {
    OBSERVED.get(rule).copied().unwrap_or(false)
}

impl grameo::builder::Builder for ExpressionBuilder<'_> {
    fn supports_choice_heads(&self) -> bool { true }
    fn supports_unobserved_failure_prefixes(&self) -> bool { true }
    fn observes_rule(&self, rule: usize) -> bool { observed(rule) }
    fn opaque_rule(&self, _: usize) -> bool { true }
    fn supports_pratt(&self, rule: usize) -> bool { is_expression_root(rule) }
    fn checkpoint(&self) -> usize { self.log.len() }
    fn rollback(&mut self, mark: usize) {
        self.log.truncate(mark);
        while self.slots.last().is_some_and(|slot| slot.event >= mark) {
            self.slots.pop();
        }
    }
    fn begin(&mut self, _: usize, _: usize) { self.log.push(Event::Begin); }
    fn finish(&mut self, rule: usize, span: Span, _: &str) {
        if is_expression_root(rule) {
            self.collapse(rule, span);
        } else {
            self.log.push(Event::Finish(rule, span));
        }
    }
    fn reduce(&mut self, _: usize, _: grameo::pratt::Fixity, span: Span) {
        self.log.push(Event::Reduce(span));
    }
    fn prune_capture_children(&self, rule: usize, span: Span) -> bool {
        matches!(self.log.last(), Some(Event::Root(id))
            if self.slots[*id].key == (rule, span.start, span.end))
    }
}

// ── Replay ──────────────────────────────────────────────────────────────────

enum Value {
    /// A constructed expression and its source span; `root` names the
    /// expression rule whose own value this is, until an enclosing construct
    /// consumes it.
    Expr { expr: Expression, span: Span, root: Option<usize> },
    /// A construct outside the owned subset.
    Incomplete,
}

enum Item {
    Frame,
    Value(Value),
}

/// Replays one expression rule's events.
struct Replay<'s> {
    source: &'s str,
    stack: Vec<Item>,
    released: Vec<(RegionKey, Expression)>,
}

/// The single owned non-root expression among `children`.
fn single_inner(children: &[Value]) -> Option<Span> {
    match children {
        [Value::Expr { span, root: None, .. }] => Some(*span),
        _ => None,
    }
}

/// Go whitespace and comments only (`WHITESPACE`, `COMMENT` in the grammar).
fn is_trivia(mut text: &str) -> bool {
    loop {
        text = text.trim_start_matches([' ', '\t', '\r', '\n']);
        if text.is_empty() {
            return true;
        }
        if let Some(rest) = text.strip_prefix("//") {
            match rest.find(['\n', '\r']) {
                Some(end) => text = &rest[end..],
                None => return true,
            }
        } else if let Some(rest) = text.strip_prefix("/*") {
            match rest.find("*/") {
                Some(end) => text = &rest[end + 2..],
                None => return false,
            }
        } else {
            return false;
        }
    }
}

fn is_go_identifier(text: &str) -> bool {
    let mut chars = text.chars();
    chars.next().is_some_and(|first| first.is_alphabetic() || first == '_')
        && chars.all(|ch| ch.is_alphanumeric() || ch == '_')
}

impl<'s> Replay<'s> {
    fn new(source: &'s str) -> Self {
        Self { source, stack: Vec::new(), released: Vec::new() }
    }

    fn event(&mut self, event: Event, slots: &mut [Slot]) {
        match event {
            Event::Begin => self.stack.push(Item::Frame),
            Event::Finish(rule, span) => self.finish(rule, span),
            Event::Reduce(span) => self.reduce(span),
            Event::Root(id) => {
                let slot = &mut slots[id];
                let (rule, start, end) = slot.key;
                let expr = slot.expr.take().expect("a root value is consumed once");
                self.stack.push(Item::Value(Value::Expr { expr, span: Span { start, end }, root: Some(rule) }));
            }
            Event::Released(id) => {
                let slot = &mut slots[id];
                let expr = slot.expr.take().expect("a released value is moved once");
                self.released.push((slot.key, expr));
            }
            Event::Incomplete => self.stack.push(Item::Value(Value::Incomplete)),
        }
    }

    fn text(&self, span: Span) -> &str {
        &self.source[span.start..span.end]
    }

    /// Keep the completed expression roots among values no construct consumes.
    fn release(&mut self, values: impl IntoIterator<Item = Value>) {
        for value in values {
            if let Value::Expr { expr, span, root: Some(rule) } = value {
                self.released.push(((rule, span.start, span.end), expr));
            }
        }
    }

    /// The values above the innermost open frame, which closes.
    fn children(&mut self) -> Vec<Value> {
        let mut values = Vec::new();
        while let Some(item) = self.stack.pop() {
            match item {
                Item::Value(value) => values.push(value),
                Item::Frame => break,
            }
        }
        values.reverse();
        values
    }

    /// The values of the collapsing expression rule itself.
    fn remaining(&mut self) -> Vec<Value> {
        let mut values = Vec::new();
        for item in std::mem::take(&mut self.stack) {
            if let Item::Value(value) = item {
                values.push(value);
            }
        }
        values
    }

    fn incomplete(&mut self, children: Vec<Value>) -> Value {
        self.release(children);
        Value::Incomplete
    }

    fn take_expr(children: Vec<Value>) -> Expression {
        match children.into_iter().next() {
            Some(Value::Expr { expr, .. }) => expr,
            _ => unreachable!("an owned child was checked"),
        }
    }

    fn root(&mut self, rule: usize, span: Span, children: Vec<Value>) -> Value {
        match single_inner(&children) {
            Some(_) => Value::Expr { expr: Self::take_expr(children), span, root: Some(rule) },
            None => self.incomplete(children),
        }
    }

    fn finish(&mut self, rule: usize, span: Span) {
        use Rule as R;
        let children = self.children();
        let value = if rule == R::operand as usize
            || rule == R::if_operand as usize
            || rule == R::range_operand as usize
        {
            self.operand(span, children)
        } else if rule == R::primary as usize
            || rule == R::if_primary as usize
            || rule == R::range_primary as usize
        {
            // `primary = operand ~ postfix_op*`: without a postfix operator
            // only trivia follows the operand within the primary's span.
            match single_inner(&children) {
                Some(inner) if is_trivia(&self.source[inner.end..span.end]) => {
                    Value::Expr { expr: Self::take_expr(children), span, root: None }
                }
                _ => self.incomplete(children),
            }
        } else if rule == R::unary_expression as usize
            || rule == R::if_unary_expression as usize
            || rule == R::range_unary_expression as usize
        {
            self.unary(span, children)
        } else {
            self.incomplete(children)
        };
        self.stack.push(Item::Value(value));
    }

    /// `operand = literal | "(" expression ")" | conversions | composite_literal
    /// | function_literal | ident_name`. A literal or identifier has no
    /// observed descendants; its first character selects the literal
    /// alternative that matched. A parenthesized expression is the operand's
    /// only root, enclosed by the parentheses and trivia.
    fn operand(&mut self, span: Span, children: Vec<Value>) -> Value {
        let text = self.text(span);
        if children.is_empty() {
            if text.ends_with([' ', '\t', '\r', '\n']) {
                return Value::Incomplete;
            }
            let mut chars = text.chars();
            let first = chars.next();
            let second = chars.next();
            let kind = match first {
                Some('0'..='9') => Some(Rule::numeric_literal),
                Some('.') if second.is_some_and(|ch| ch.is_ascii_digit()) => Some(Rule::numeric_literal),
                Some('"' | '`') => Some(Rule::string_literal),
                Some('\'') => Some(Rule::rune_literal),
                _ => None,
            };
            let expr = match kind {
                Some(kind) => go_literal_value(kind, text).unwrap_or_else(Expression::null),
                None if is_go_identifier(text) => go_operand_ident(text),
                None => return Value::Incomplete,
            };
            return Value::Expr { expr, span, root: None };
        }
        if let [Value::Expr { span: inner, root: Some(_), .. }] = children.as_slice() {
            let before = &self.source[span.start..inner.start];
            let after = &self.source[inner.end..span.end];
            if before.strip_prefix('(').is_some_and(is_trivia)
                && after.trim_end_matches([' ', '\t', '\r', '\n']).strip_suffix(')').is_some_and(is_trivia)
            {
                return Value::Expr { expr: Self::take_expr(children), span, root: None };
            }
        }
        self.incomplete(children)
    }

    /// `unary_expression = unary_op ~ unary_expression | primary`; the
    /// operator is the source between the two spans.
    fn unary(&mut self, span: Span, children: Vec<Value>) -> Value {
        let Some(inner) = single_inner(&children) else { return self.incomplete(children); };
        if inner.start == span.start {
            return Value::Expr { expr: Self::take_expr(children), span, root: None };
        }
        let op = self.source[span.start..inner.start].trim();
        if !matches!(op, "-" | "+" | "!" | "^" | "*" | "&" | "<-") {
            return self.incomplete(children);
        }
        let op = op.to_string();
        Value::Expr { expr: go_unary_expr(&op, Self::take_expr(children)), span, root: None }
    }

    /// A Pratt reduction combines the two values on top of the current frame.
    fn reduce(&mut self, span: Span) {
        let operands = matches!(self.stack.as_slice(),
            [.., Item::Value(Value::Expr { root: None, .. }), Item::Value(Value::Expr { root: None, .. })]);
        if operands {
            let Some(Item::Value(Value::Expr { expr: right, span: right_span, .. })) = self.stack.pop()
                else { unreachable!() };
            let Some(Item::Value(Value::Expr { expr: left, span: left_span, .. })) = self.stack.pop()
                else { unreachable!() };
            let op = self.text(span).trim().to_string();
            self.stack.push(Item::Value(Value::Expr {
                expr: build_go_binary_expr(&op, left, right),
                span: Span { start: left_span.start, end: right_span.end },
                root: None,
            }));
            return;
        }
        // Operands that are not owned values: the region is incomplete.
        let mut taken = Vec::new();
        for _ in 0..2 {
            match self.stack.pop() {
                Some(Item::Value(value)) => taken.push(value),
                Some(frame @ Item::Frame) => {
                    self.stack.push(frame);
                    break;
                }
                None => break,
            }
        }
        self.release(taken);
        self.stack.push(Item::Value(Value::Incomplete));
    }
}

// ── Shared construction (walker and builder) ────────────────────────────────

/// The value of a Go literal token of the given rule, or `None` when the
/// token's text is not representable.
pub(crate) fn go_literal_value(rule: Rule, text: &str) -> Option<Expression> {
    match rule {
        Rule::numeric_literal => go_numeric_literal(text),
        Rule::string_literal => Some(Expression::new(ExprKind::Lit(Literal::Str(
            crate::walker::unquote(text),
        )))),
        Rule::bool_literal => Some(Expression::new(ExprKind::Lit(Literal::Bool(text == "true")))),
        Rule::nil_literal => Some(Expression::new(ExprKind::Lit(Literal::Null))),
        Rule::rune_literal => {
            let rune = crate::walker::unquote(text);
            let code = rune.chars().next().map(|ch| ch as i64).unwrap_or(0);
            Some(Expression::new(ExprKind::Lit(Literal::Int(code))))
        }
        _ => None,
    }
}

/// Go integer, floating-point and imaginary literals (spec "Integer
/// literals", "Floating-point literals", "Imaginary literals"). An integer
/// literal is an exact untyped constant of any size: see [`go_exact_integer`].
fn go_numeric_literal(text: &str) -> Option<Expression> {
    let mut s = text.replace('_', "");
    let imaginary = s.ends_with('i');
    if imaginary {
        s.pop();
    }
    let hex = s.starts_with("0x") || s.starts_with("0X");
    let float = if hex { s.contains(['p', 'P']) } else { s.contains(['.', 'e', 'E']) };
    let value = if float {
        if hex { go_hex_float(&s[2..]).map(Expression::float) } else { s.parse::<f64>().ok().map(Expression::float) }
    } else if hex {
        go_exact_integer(u128::from_str_radix(&s[2..], 16).ok()?)
    } else if s.starts_with("0b") || s.starts_with("0B") {
        go_exact_integer(u128::from_str_radix(&s[2..], 2).ok()?)
    } else if s.starts_with("0o") || s.starts_with("0O") {
        go_exact_integer(u128::from_str_radix(&s[2..], 8).ok()?)
    } else if s.len() > 1 && s.starts_with('0') {
        // A leading `0` without a base prefix is octal (`0777` is 511).
        go_exact_integer(u128::from_str_radix(&s[1..], 8).ok()?)
    } else {
        go_exact_integer(s.parse::<u128>().ok()?)
    }?;
    if imaginary {
        return Some(crate::walker::go_complex_value_expr(Expression::int(0), value));
    }
    Some(value)
}

/// An exact integer constant. A value beyond `i64` is the sum
/// `high * 2^32 + low` of exact parts: constant folding recovers the value
/// exactly, and float evaluation rounds it once to the nearest float.
fn go_exact_integer(value: u128) -> Option<Expression> {
    if let Ok(small) = i64::try_from(value) {
        return Some(Expression::int(small));
    }
    let high = go_exact_integer(value >> 32)?;
    Some(Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(high),
            right: Box::new(Expression::int(1 << 32)),
        })),
        right: Box::new(Expression::int((value & 0xffff_ffff) as i64)),
    }))
}

/// The value of a hexadecimal floating-point mantissa and exponent
/// (`1.8p3` after the `0x` prefix).
fn go_hex_float(text: &str) -> Option<f64> {
    let (mantissa, exponent) = text.split_once(['p', 'P'])?;
    let exponent: i32 = exponent.parse().ok()?;
    let (whole, fraction) = mantissa.split_once('.').unwrap_or((mantissa, ""));
    let digits = format!("{whole}{fraction}");
    let value = u128::from_str_radix(if digits.is_empty() { "0" } else { &digits }, 16).ok()? as f64;
    Some(value * 2f64.powi(exponent - 4 * fraction.len() as i32))
}

/// An identifier operand; the predeclared `nil`, `true` and `false` are values.
pub(crate) fn go_operand_ident(name: &str) -> Expression {
    match name {
        "nil" => Expression::new(ExprKind::Lit(Literal::Null)),
        "true" => Expression::new(ExprKind::Lit(Literal::Bool(true))),
        "false" => Expression::new(ExprKind::Lit(Literal::Bool(false))),
        _ => Expression::new(ExprKind::Ident(name.to_string())),
    }
}

/// A Go unary operator applied to its operand; `<-` is a channel receive.
pub(crate) fn go_unary_expr(op: &str, operand: Expression) -> Expression {
    let op = match op {
        "-" => UnaryOp::Neg,
        "!" => UnaryOp::Not,
        "+" => UnaryOp::Pos,
        "^" => UnaryOp::BitNot,
        "*" => UnaryOp::Deref,
        "&" => UnaryOp::AddrOf,
        "<-" => return crate::walker::chan_recv(operand),
        _ => UnaryOp::Pos,
    };
    Expression::new(ExprKind::Unary { op, expr: Box::new(operand) })
}

fn go_bin_op(op: &str) -> BinOp {
    match op {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        "<=" => BinOp::LtEq,
        ">" => BinOp::Gt,
        ">=" => BinOp::GtEq,
        "&&" => BinOp::And,
        "||" => BinOp::Or,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "&^" => BinOp::BitAnd,
        _ => BinOp::Add,
    }
}

/// A Go binary operation; `a &^ b` (bit clear) is `a & ^b`.
pub(crate) fn build_go_binary_expr(op: &str, left: Expression, right: Expression) -> Expression {
    if op == "&^" {
        Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(left),
            right: Box::new(Expression::new(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new(right),
            })),
        })
    } else {
        Expression::new(ExprKind::Binary {
            op: go_bin_op(op),
            left: Box::new(left),
            right: Box::new(right),
        })
    }
}
