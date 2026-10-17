//! Incremental common-AST construction at existing PEG rule boundaries.
//! Context-dependent expressions retain captures for the contextual walker.
use super::*;
use grameo::{
    arena::{Id, MetadataArena},
    builder::Builder,
};
use std::collections::HashMap;

pub(super) type ExpressionCache = HashMap<
    (usize, usize), CachedExpression, vybe_runtime::chunk::FxBuildHasher,
>;

#[cfg(test)]
std::thread_local! {static ELIGIBILITY_CENSUS: std::cell::Cell<(usize,usize)> = const { std::cell::Cell::new((0,0)) };}
#[cfg(test)]
std::thread_local! {static CONTEXT_CENSUS:std::cell::Cell<[u64;7]>=const {std::cell::Cell::new([0;7])};}
#[cfg(test)]
fn census_add(index:usize,count:u64){CONTEXT_CENSUS.with(|cell|{let mut counters=cell.get();counters[index]+=count;cell.set(counters);});}
struct Metadata {
    raw: grameo::Span,
    anchor: Option<grameo::Span>,
    pratt_ops: u8,
    normalization: u8,
}
const NEEDS_GLOBAL_BINDING: u8 = 1;
const NEEDS_ARITHMETIC_FOLD: u8 = 2;
const NEEDS_CONTEXTUAL_BINARY: u8 = 4;
const NEEDS_SUPERGLOBAL_BINDING: u8 = 8;

fn leaf_normalization(kind: &ExprKind) -> u8 {
    // Conservative and constant-time: the current scope decides whether
    // identifiers need binding. Avoid a name-table lookup during matching.
    if let ExprKind::Ident(name) = kind {
        NEEDS_GLOBAL_BINDING | if PHP_SUPERGLOBAL_NAMES.contains(&name.as_str()) { NEEDS_SUPERGLOBAL_BINDING } else { 0 }
    } else { 0 }
}
fn binary_normalization(op: BinOp, left: &Expression, right: &Expression) -> u8 {
    if matches!(op, BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq | BinOp::Eq | BinOp::NotEq | BinOp::StrictEq | BinOp::StrictNotEq | BinOp::BitAnd | BinOp::BitXor | BinOp::BitOr | BinOp::And | BinOp::Or) {
        return NEEDS_CONTEXTUAL_BINARY;
    }
    let fold = if op == BinOp::Div {
        php_expr_is_literal_zero(right)
    } else if matches!(op, BinOp::Add | BinOp::Sub) {
        match (php_literal_or_php_int_const_i64(left), php_literal_or_php_int_const_i64(right)) {
            (Some(left), Some(right)) => if op == BinOp::Add {
                left.checked_add(right).is_none()
            } else { left.checked_sub(right).is_none() },
            _ => false,
        }
    } else { false };
    if fold { NEEDS_ARITHMETIC_FOLD } else { 0 }
}
enum MetadataUndo {
    RecipeOperands(Metadata, Metadata, bool, bool),
    Anchor(Option<grameo::Span>),
    Capture(Option<grameo::tree::CaptureHandle>),
    Operands(Metadata, Metadata),
    ExpressionSpan(Span),
    SourceState(SourceState),
    SourceFrame(SourceFrame),
}
pub(super) enum ContextRecipe {
    Owned(Expression),
    Atom { rule: Rule, raw: grameo::Span, capture: Option<grameo::tree::CaptureHandle> },
    Binary { op: BinOp, span: Span, left: Box<ContextRecipe>, right: Box<ContextRecipe> },
}
#[derive(Clone, Copy)]
struct SourceState { head: Option<Id>, deferred: usize, last_intent: Option<Id> }
#[derive(Clone, Copy)]
struct SourceFrame { parent: Option<Id>, start: usize, selector: Option<usize>, target: Option<grameo::Span>, targets: usize }
struct SourceHooks { state: Id, bindings: Vec<grameo::source_units::Binding<'static>>, root: usize, target: usize, observed: Vec<usize>, deferred: [usize;4] }
enum Value {
    SourceState(SourceState),
    SourceFrame(SourceFrame),
    SourceIntent { request: grameo::source_units::Request<String>, previous: Option<Id>, creation_mark: usize },
    Declaration(Option<grameo::tree::CaptureHandle>),
    Recipe { recipe: ContextRecipe, metadata: Metadata },
    Marker(Rule),
    Ready {
        expression: Expression,
        metadata: Metadata,
    },
    Operator(BinOp),
}
pub(super) enum CachedExpression {
    Contextual(Option<ContextRecipe>),
    String {
        first: Option<String>,
        raw: grameo::Span,
        span: Span,
    },
    Ident {
        first: Option<String>,
        raw: grameo::Span,
        span: Span,
    },
    Number {
        first: Option<Literal>,
        raw: grameo::Span,
        span: Span,
    },
    Pending(Option<Expression>),
    // First consumption moves the AST; repeated inspection recaptures PEG.
    Deferred { first: Option<Expression>, normalization: u8 },
    Owned(Vec<Replay>),
}
pub(super) enum Replay {
    Leaf { raw: grameo::Span, span: Span },
    Binary { op: BinOp, span: Span },
    // Unsupported source leaves retain only their original value, not a copy
    // of the enclosing arithmetic tree. Scope binding occurs on each replay.
    Value(Box<Expression>),
}
impl Value {
    fn metadata(&self) -> Option<&Metadata> { match self { Self::Ready {metadata,..} | Self::Recipe {metadata,..} => Some(metadata), _=>None } }
    fn expression_span_mut(&mut self) -> &mut Span { match self {Self::Ready {expression,..}=> &mut expression.span,Self::Recipe {recipe:ContextRecipe::Binary {span,..},..}=>span,_=>unreachable!()} }
    fn into_recipe(self) -> (ContextRecipe,Metadata,bool) {match self {Self::Ready {expression,metadata}=>(ContextRecipe::Owned(expression),metadata,true),Self::Recipe {recipe,metadata}=>(recipe,metadata,false),_=>unreachable!()} }
}
impl CachedExpression {
    fn new(expression: Expression, raw: grameo::Span, _source: &str, _lines: &[usize]) -> Self {
        match expression {
            Expression { kind: ExprKind::Lit(Literal::Str(value)), span } => Self::String {
                first: Some(value), raw, span,
            },
            Expression {
                kind: ExprKind::Ident(name),
                span,
            } => Self::Ident {
                first: Some(name),
                raw,
                span,
            },
            Expression {
                kind: ExprKind::Lit(value),
                span,
            } => Self::Number {
                first: Some(value),
                raw,
                span,
            },
            other => Self::Pending(Some(other)),
        }
    }
    fn bind(&mut self, input: &str, module_reads: bool, superglobal_reads: bool, lines: &[usize]) -> Expression {
        if let Self::Deferred { first, normalization } = self {
            let expression = first.take().expect("first deferred expression consumption");
            if *normalization & NEEDS_ARITHMETIC_FOLD == 0
                && (!module_reads || *normalization & NEEDS_GLOBAL_BINDING == 0)
                && (!superglobal_reads || *normalization & NEEDS_SUPERGLOBAL_BINDING == 0) {
                return expression;
            }
            return transform_owned_impl(expression, Some(module_reads), superglobal_reads, None).0;
        }
        // Captures can retain values that contextual lowering never consumes.
        // Only materialize their normalization and replay state on first use.
        if let Self::Pending(first) = self {
            let expression = first.take().expect("pending expression is owned");
            let (expression, _, replay) = transform_owned_impl(
                expression, Some(module_reads), superglobal_reads, Some((input, lines)),
            );
            *self = Self::Owned(replay);
            return expression;
        }
        let expression = match self {
            Self::Contextual(_) => unreachable!("contextual recipe requires the live walker"),
            Self::Pending(_) => unreachable!("pending expression was initialized"),
            Self::Deferred { .. } => unreachable!("deferred expression was consumed"),
            Self::String { first, raw, span } => Expression::with_span(
                ExprKind::Lit(Literal::Str(first.take().unwrap_or_else(|| {
                    let raw = &input[raw.start..raw.end];
                    let body = &raw[1..raw.len() - 1];
                    if raw.starts_with('\'') { decode_php_single_quoted_literal(body) }
                    else { unmask_php_literal_tags_owned(decode_php_double_quoted_literal(body)) }
                }))), *span,
            ),
            Self::Ident { first, raw, span } => Expression::with_span(
                ExprKind::Ident(
                    first
                        .take()
                        .unwrap_or_else(|| strip_dollar(&input[raw.start..raw.end]).to_string()),
                ),
                *span,
            ),
            Self::Number { first, raw, span } => Expression::with_span(
                ExprKind::Lit(first.take().unwrap_or_else(|| {
                    let ExprKind::Lit(value) =
                        numeric_literal::number(&input[raw.start..raw.end]).kind
                    else {
                        unreachable!("cached numeric literal")
                    };
                    value
                })),
                *span,
            ),
            Self::Owned(replay) => return replay_expression(replay, input, module_reads, superglobal_reads),
        };
        bind_owned_leaf(expression, module_reads, superglobal_reads)
    }
}

enum ModeChange { Push, Pop(bool) }
struct ModeUndo {
    arena_mark: usize,
    change: ModeChange,
}

#[derive(Default)]
pub(super) struct Ast<'i> {
    source: &'i str,
    committed_pratt_values: usize,
    unit_commit: Option<&'i std::cell::Cell<Option<usize>>>,
    source_hooks: Option<SourceHooks>,
    source_intents: Vec<grameo::source_units::Request<String>>,
    pratt_only: bool,
    // Adjacent nested Pratt boundaries ask about the same immutable source
    // offset. Cache only that pure admission decision, never parser/AST state.
    last_pratt_eligibility: Option<(usize, bool, bool)>,
    #[cfg(feature = "service-verification")]
    admission_cache_enabled: bool,
    #[cfg(feature = "service-verification")]
    admission_census_builder: usize,
    modes: Vec<bool>,
    mode_undo: Vec<ModeUndo>,
    arena: MetadataArena<Value, MetadataUndo>,
    lines: Vec<usize>,
    markers: Vec<Option<Id>>,
    hooks: Vec<bool>,
    operators: Vec<Id>,
    expressions: ExpressionCache,
    declarations: Vec<Option<grameo::tree::CaptureHandle>>,
}
fn selected(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::logic_or_expression | Rule::bit_or_expression | Rule::equality_expression | Rule::logic_or_assignment_operand | Rule::logic_and_assignment_operand | Rule::equality_assignment_operand | Rule::comparison_expression
            | Rule::concat_expression
            | Rule::shift_expression
            | Rule::additive_expression
            | Rule::multiplicative_expression
            | Rule::boolean_not_expression
            | Rule::instanceof_expression
            | Rule::power_expression
            | Rule::unary_expression
            | Rule::postfix_expression
            | Rule::primary_expression
    )
}
fn owned_operator(op: BinOp) -> bool {
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
            | BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq | BinOp::Eq | BinOp::NotEq | BinOp::StrictEq | BinOp::StrictNotEq | BinOp::BitAnd | BinOp::BitXor | BinOp::BitOr | BinOp::And | BinOp::Or
    )
}
impl<'i> Ast<'i> {
    pub(super) fn new(source: &'i str) -> Self {
        Self::new_with_lines(source, build_line_starts(source))
    }
    fn new_with_lines(source: &'i str, lines: Vec<usize>) -> Self {
        let rules = [
            Rule::logic_or_expression, Rule::logic_or_assignment_operand, Rule::logic_and_assignment_operand, Rule::bit_or_expression, Rule::equality_expression, Rule::equality_assignment_operand, Rule::comparison_expression,
            Rule::concat_expression,
            Rule::shift_expression,
            Rule::additive_expression,
            Rule::multiplicative_expression,
            Rule::boolean_not_expression,
            Rule::instanceof_expression,
            Rule::power_expression,
            Rule::unary_expression,
            Rule::postfix_expression,
            Rule::primary_expression,
        ];
        let mut builder = Self {
            source,
            #[cfg(feature = "service-verification")]
            admission_cache_enabled: crate::PHP_PRATT_ADMISSION_CACHE_FOR_VALIDATION.with(|flag| flag.get()),
            #[cfg(feature = "service-verification")]
            admission_census_builder: crate::admission_census_builder_for_validation(),
            lines,
            markers: vec![None; rules.iter().map(|rule| *rule as usize).max().unwrap() + 1],
            ..Self::default()
        };
        // Immutable sentinels can be shared across nested frames. Allocate
        // once, before the parser's initial rollback cursor, rather than one
        // full AST-sized arena slot per transparent rule invocation.
        for rule in rules {
            builder.markers[rule as usize] = Some(builder.arena.alloc(Value::Marker(rule)));
        }
        // Operators are immutable semantic tokens too. Reuse their handles
        // instead of allocating an AST-sized arena slot for each occurrence.
        for op in [
            BinOp::Add,
            BinOp::Sub,
            BinOp::Mul,
            BinOp::Div,
            BinOp::Mod,
            BinOp::Shl,
            BinOp::Shr,
            BinOp::Concat,
            BinOp::Lt,
            BinOp::Gt,
            BinOp::LtEq,
            BinOp::GtEq,
            BinOp::Spaceship,
        ] {
            builder
                .operators
                .push(builder.arena.alloc(Value::Operator(op)));
        }
        builder.hooks = builder.markers.iter().map(Option::is_some).collect();
        for rule in [
            Rule::number_lit,
            Rule::simple_variable,
            Rule::sq_string,
            Rule::dq_string,
            Rule::mul_op,
            Rule::additive_op,
            Rule::shift_op,
            Rule::concat_op,
            Rule::cmp_op,
        ] {
            if builder.hooks.len() <= rule as usize {
                builder.hooks.resize(rule as usize + 1, false);
            }
            builder.hooks[rule as usize] = true;
        }
        builder
    }

    pub(super) fn pratt_only(source: &'i str) -> Self {
        let mut builder = Self::new(source);
        builder.pratt_only = true;
        builder.arena.commit();
        builder
    }
    pub(super) fn with_unit_commit(mut self, marker: &'i std::cell::Cell<Option<usize>>) -> Self {
        self.unit_commit = Some(marker);
        self
    }
    fn mode_enabled(&self) -> bool {
        !self.pratt_only || self.modes.last().copied().unwrap_or(false)
    }
    fn push_mode(&mut self, enabled: bool) {
        self.mode_undo.push(ModeUndo {
            arena_mark: self.arena.checkpoint(), change: ModeChange::Push,
        });
        self.modes.push(enabled);
    }
    fn pop_mode(&mut self) {
        self.mode_undo.push(ModeUndo {
            arena_mark: self.arena.checkpoint(),
            change: ModeChange::Pop(self.modes.pop().expect("active expression mode")),
        });
    }
    fn rollback_mark(&mut self, mark: usize) {
        // Both journals contribute one step per operation. A mode event's
        // position is its arena cursor plus preceding mode events, so marks
        // need no packed bit fields or per-checkpoint allocation (also on Wasm).
        while self.mode_undo.last().is_some_and(|undo| {
            undo.arena_mark + self.mode_undo.len() - 1 >= mark
        }) {
            match self.mode_undo.pop().unwrap().change {
                ModeChange::Push => { self.modes.pop().expect("mode rollback"); }
                ModeChange::Pop(enabled) => self.modes.push(enabled),
            }
        }
        self.arena.rollback(mark - self.mode_undo.len());
    }
    fn span(&self, span: grameo::Span) -> Span {
        let (start_line, start_col) = offset_to_line_col(span.start, &self.lines);
        let (end_line, end_col) = offset_to_line_col(span.end, &self.lines);
        Span {
            start_line,
            start_col,
            end_line,
            end_col,
        }
    }
    fn push(&mut self, value: Value) {
        let id = self.arena.alloc(value);
        self.arena.push(id);
    }
    fn anchor(&mut self, id: Id, span: grameo::Span) {
        self.set_anchor(id, span);
        self.arena.push(id);
    }
    fn set_anchor(&mut self, id: Id, span: grameo::Span) {
        if matches!(self.arena.get(id), Value::Ready { metadata: Metadata { anchor: Some(anchor), .. }, .. } if *anchor == span)
        {
            return;
        }
        self.arena.edit(
            id,
            |value| match value {
                Value::Ready { metadata, .. } | Value::Recipe { metadata, .. } => {
                    MetadataUndo::Anchor(metadata.anchor.replace(span))
                }
                _ => unreachable!(),
            },
            |value, previous| match value {
                Value::Ready { metadata, .. } | Value::Recipe { metadata, .. } => {
                    metadata.anchor = match previous {
                        MetadataUndo::Anchor(anchor) => anchor,
                        _ => unreachable!(),
                    }
                }
                _ => unreachable!(),
            },
        );
    }
    /// Seal only the completed top-level statement's expression recipes, then
    /// start a fresh local arena with the same full-source line table. The
    /// producer owns the outer generation/abort journal; no sentinel Id escapes.
    pub(super) fn drain_statement_boundary(mut self, consumed: usize)
        -> (Self, ExpressionCache, usize) {
        let source = self.source;
        self.committed(source, consumed);
        let pratt = self.committed_pratt_values;
        let expressions = std::mem::take(&mut self.expressions);
        let lines = std::mem::take(&mut self.lines);
        (Self::new_with_lines(source, lines), expressions, pratt)
    }
    pub(super) fn finish(self) -> (ExpressionCache, Vec<usize>, Vec<Option<grameo::tree::CaptureHandle>>) {
        (self.expressions, self.lines, self.declarations)
    }
}
impl Builder for Ast<'_> {
    fn supports_choice_heads(&self) -> bool { true }
    // Only generated zero-consumption failure proofs whose reached hooks are
    // all unobserved may bypass matching. Observed scope/AST hooks stay intact.
    fn supports_unobserved_failure_prefixes(&self) -> bool { true }
    fn committed(&mut self, _: &str, consumed: usize) {
        // This notification is opt-in and runs only after Grameo validated the
        // selected entry. It does not publish AST or execute source requests.
        if let Some(marker) = self.unit_commit { marker.set(Some(consumed)); }
        self.expressions = std::mem::take(&mut self.arena)
            .into_live_nodes()
            .filter_map(|value| match value {
                Value::SourceIntent { request, .. } => { self.source_intents.push(request); None },
                Value::Declaration(capture) => { self.declarations.push(capture); None },
                // A lone contextual atom already retains its complete capture.
                // Only operator composition needs a deferred contextual root.
                Value::Recipe { recipe: ContextRecipe::Atom {..}, .. } => None,
                Value::Recipe { mut recipe, metadata: Metadata { anchor: Some(span), pratt_ops, .. } } => {
                    if pratt_ops > 0 { self.committed_pratt_values += 1; }
                    if pratt_ops > 1 { normalize_recipe_spans(&mut recipe, self.source, &self.lines); }
                    Some(((span.start,span.end), CachedExpression::Contextual(Some(recipe))))
                },
                Value::Ready {
                    mut expression,
                    metadata:
                        Metadata {
                            raw,
                            anchor: Some(span),
                            pratt_ops,
                            normalization,
                            ..
                        },
                } => {
                    if pratt_ops > 0 { self.committed_pratt_values += 1; }
                    if pratt_ops > 1 {
                        normalize_pratt_spans(&mut expression, self.source, &self.lines);
                    }
                    let cached = match CachedExpression::new(expression, raw, self.source, &self.lines) {
                        CachedExpression::Pending(first) => CachedExpression::Deferred { first, normalization },
                        other => other,
                    };
                    Some(((span.start, span.end), cached))
                },
                _ => None,
            })
            .collect();
        #[cfg(test)] {
            census_add(0,self.expressions.len() as u64);
            for cached in self.expressions.values(){
                if let CachedExpression::Contextual(Some(recipe))=cached {
                    census_add(1,1);let mut pending=vec![recipe];
                    while let Some(recipe)=pending.pop(){match recipe {
                        ContextRecipe::Binary {left,right,..}=>{census_add(2,1);pending.push(left);pending.push(right);},
                        ContextRecipe::Atom {..}=>census_add(3,1),
                        ContextRecipe::Owned(_)=>census_add(4,1),
                    }}
                }
            }
        }
    }
    fn observes_capture(&self, rule: usize) -> bool {
        declaration_rule(rule) || matches!(Rule::from_capture(grameo::CaptureRule::User(rule)),
            Some(Rule::boolean_not_expression | Rule::logic_or_assignment_operand
                | Rule::logic_and_assignment_operand | Rule::equality_assignment_operand))
    }
    fn try_finish_capture(&mut self, rule: usize, span: grameo::Span, source: &str, capture: Option<grameo::tree::CaptureHandle>) -> Result<(), grameo::builder::BuildError> {
        self.source_finish(rule, span, source)?;
        self.finish_ast_rule(rule, span, source);
        if declaration_rule(rule) { self.arena.alloc(Value::Declaration(capture)); }
        if let Some(&id) = self.arena.values().last()
            && matches!(self.arena.get(id), Value::Recipe { recipe: ContextRecipe::Atom {rule: atom_rule, raw, ..}, .. }
                if *atom_rule as usize == rule && *raw == span) {
            self.arena.edit(id, |value| {
                let Value::Recipe {recipe: ContextRecipe::Atom {capture: old,..},..} = value else {unreachable!()};
                MetadataUndo::Capture(std::mem::replace(old, capture))
            }, |value, undo| {
                let Value::Recipe {recipe: ContextRecipe::Atom {capture,..},..} = value else {unreachable!()};
                let MetadataUndo::Capture(old) = undo else {unreachable!()};
                *capture = old;
            });
        }
        Ok(())
    }
    fn observes_rule(&self, rule: usize) -> bool {
        if self.source_hooks.as_ref().is_some_and(|hooks| hooks.observed.binary_search(&rule).is_ok()) {return true;}
        if declaration_rule(rule) { return true; }
        if rule == Rule::comparison_expression as usize || rule == Rule::bit_or_expression as usize || rule == Rule::logic_or_expression as usize { return true; }
        if !self.pratt_only {
            // Pratt supplies operators directly; only atom boundaries and the
            // lexical leaves they can own need builder frames in this route.
            return matches!(Rule::from_capture(grameo::CaptureRule::User(rule)), Some(
                Rule::logic_or_assignment_operand | Rule::logic_and_assignment_operand | Rule::equality_assignment_operand | Rule::boolean_not_expression | Rule::number_lit | Rule::simple_variable | Rule::sq_string | Rule::dq_string
            ));
        }
        self.hooks.get(rule).copied().unwrap_or(false) && self.mode_enabled()
    }

    fn supports_pratt_at(&self, rule: usize, _offset: usize, _source: &str) -> bool {
        if Rule::from_capture(grameo::CaptureRule::User(rule)) .is_none_or(|rule|!matches!(rule,Rule::comparison_expression | Rule::bit_or_expression | Rule::logic_or_expression)) { return false; }
        self.mode_enabled()
    }

    fn accepts_pratt(&self, _: usize, span: grameo::Span) -> bool {
        if !self.pratt_only && self.modes.last().copied().unwrap_or(false)
            && self.arena.values().last().is_some_and(|id|matches!(self.arena.get(*id),Value::Recipe {..})) {
            // A source hint may optimistically select capture suppression.
            // Deferred operands require captures, so recapture PEG instead.
            return false;
        }
        self.arena.values().last().and_then(|id|self.arena.get(*id).metadata()).is_some_and(|metadata| {
            let raw=metadata.raw;
            raw.start == span.start && raw.end <= span.end && self.source[raw.end..span.end].trim().is_empty()
        })
    }

    fn suppress_pratt_captures(&self, rule: usize) -> bool {
        // Complete islands already replace descendant captures with an owned
        // contextual replay value. Incomplete construction recaptures PEG.
        (rule == Rule::comparison_expression as usize || rule == Rule::bit_or_expression as usize || rule == Rule::logic_or_expression as usize) && (self.pratt_only || self.modes.last().copied().unwrap_or(false))
    }

    fn reduce(&mut self, rule: usize, _: grameo::pratt::Fixity, operator: grameo::Span) {
        // Guarded, non-atomic bitwise rules include implicit trailing trivia.
        // The binding's rule identity carries the operator independently of it.
        let op = match Rule::from_capture(grameo::CaptureRule::User(rule)) {
            Some(Rule::bit_and_op) => BinOp::BitAnd,
            Some(Rule::bit_xor_op) => BinOp::BitXor,
            Some(Rule::bit_or_op) => BinOp::BitOr,
            _ => parse_binop(&self.source[operator.start..operator.end]),
        };
        if !owned_operator(op) { return; }
        let values = self.arena.values();
        if values.len() < 2 { return; }
        let (left, right) = (values[values.len() - 2], values[values.len() - 1]);
        let (Some(left_metadata), Some(right_metadata))
            = (self.arena.get(left).metadata(), self.arena.get(right).metadata()) else { return; };
        // Unsupported atoms leave older values on the transactional stack.
        // Never combine across an unrepresented operand or operator.
        if left_metadata.raw.end > operator.start || operator.end > right_metadata.raw.start
            || recipe_trivia_end(left_metadata.raw.end,self.source)!=operator.start
            || recipe_trivia_end(operator.end,self.source)!=right_metadata.raw.start
        { return; }
        let raw = grameo::Span { start: left_metadata.raw.start, end: right_metadata.raw.end };
        let span = self.span(raw);
        let contextual = matches!(self.arena.get(left),Value::Recipe {..}) || matches!(self.arena.get(right),Value::Recipe {..})
            || self.source[left_metadata.raw.end..right_metadata.raw.start].trim()!=self.source[operator.start..operator.end].trim();
        self.arena.pop();
        self.arena.pop();
        if contextual {
            let value=self.arena.binary_with_metadata(left,right,|left,right| {
                let (left,left_metadata,left_ready)=left.into_recipe();let (right,right_metadata,right_ready)=right.into_recipe();
                let pratt_ops=left_metadata.pratt_ops.saturating_add(right_metadata.pratt_ops).saturating_add(1);
                (Value::Recipe {recipe:ContextRecipe::Binary {op,span,left:Box::new(left),right:Box::new(right)}, metadata:Metadata {raw,anchor:None,pratt_ops,normalization:0}},MetadataUndo::RecipeOperands(left_metadata,right_metadata,left_ready,right_ready))
            },split_binary);
            self.arena.push(value);return;
        }
        let value = self.arena.binary_with_metadata(left, right, |left, right| {
            let (Value::Ready { expression: left, metadata: left_metadata }, Value::Ready { expression: right, metadata: right_metadata })
                = (left, right) else { unreachable!() };
            let pratt_ops = left_metadata.pratt_ops.saturating_add(right_metadata.pratt_ops).saturating_add(1);
            let normalization = left_metadata.normalization | right_metadata.normalization
                | binary_normalization(op, &left, &right);
            (Value::Ready {
                expression: Expression::with_span(ExprKind::Binary { op, left: Box::new(left), right: Box::new(right) }, span),
                metadata: Metadata { raw, anchor: None, pratt_ops, normalization },
            }, MetadataUndo::Operands(left_metadata, right_metadata))
        }, split_binary);
        self.arena.push(value);
    }

    fn checkpoint(&self) -> usize {
        self.arena.checkpoint() + self.mode_undo.len()
    }
    fn rollback(&mut self, mark: usize) {
        self.rollback_mark(mark);
    }
    fn rollback_pratt(&mut self, mark: usize, _: usize, span: grameo::Span) {
        self.rollback_source_pratt(mark, span);
    }
    fn opaque_rule(&self, _: usize) -> bool {
        // Certified atomic bodies in this binding are lexical tokens. Only
        // their own finish hook matters; context is built by non-atomic rules.
        true
    }
    fn begin(&mut self, rule: usize, offset: usize) {
        self.source_begin(rule, offset);
        if rule == Rule::comparison_expression as usize || rule == Rule::bit_or_expression as usize || rule == Rule::logic_or_expression as usize {
            let enabled = self.cached_pratt_eligible(offset);
            self.push_mode(enabled);
        }
        if !self.mode_enabled() { return; }
        if let Some(Some(marker)) = self.markers.get(rule) {
            self.arena.push(*marker);
        }
    }
    fn finish(&mut self, rule: usize, span: grameo::Span, source: &str) {
        self.source_finish(rule, span, source).expect("use fallible hook for source requests");
        self.finish_ast_rule(rule, span, source);
    }
    fn try_finish(&mut self, rule:usize, span:grameo::Span, source:&str)->Result<(),grameo::builder::BuildError>{
        self.source_finish(rule,span,source)?; self.finish_ast_rule(rule,span,source);Ok(())
    }
    fn prune_capture_children(&self, rule: usize, span: grameo::Span) -> bool {
        matches!(Rule::from_capture(grameo::CaptureRule::User(rule)), Some(Rule::comparison_expression | Rule::concat_expression | Rule::shift_expression | Rule::additive_expression | Rule::multiplicative_expression))
            && self.arena.values().last().is_some_and(|id| matches!(self.arena.get(*id), Value::Ready { metadata: Metadata { anchor: Some(anchor), .. }, .. } if *anchor == span))
    }
}
impl Ast<'_> {
    fn cached_pratt_eligible(&mut self, offset: usize) -> bool {
        #[cfg(feature = "service-verification")]
        if !self.admission_cache_enabled {
            crate::record_pratt_admission_for_validation(self.admission_census_builder, offset, self.pratt_only, false);
            return self.pratt_eligible(offset, self.source);
        }
        if let Some((previous, pratt_only, eligible)) = self.last_pratt_eligibility {
            if previous == offset && pratt_only == self.pratt_only {
                #[cfg(test)]
                ELIGIBILITY_CENSUS.with(|c| { let (scans, hits) = c.get(); c.set((scans, hits + 1)); });
                #[cfg(feature = "service-verification")]
                crate::record_pratt_admission_for_validation(self.admission_census_builder, offset, self.pratt_only, true);
                return eligible;
            }
        }
        #[cfg(feature = "service-verification")]
        crate::record_pratt_admission_for_validation(self.admission_census_builder, offset, self.pratt_only, false);
        let eligible = self.pratt_eligible(offset, self.source);
        self.last_pratt_eligibility = Some((offset, self.pratt_only, eligible));
        #[cfg(test)]
        ELIGIBILITY_CENSUS.with(|c| { let (scans, hits) = c.get(); c.set((scans + 1, hits)); });
        eligible
    }
    pub(super) fn pratt_eligible(&self, offset: usize, source: &str) -> bool {
        // Avoid constructing speculative ASTs for complex operands.
        // This is conservative backend selection; PEG still validates tokens.
        let bytes = source.as_bytes();
        let leaf_end = |mut end: usize| -> Option<usize> {
            match bytes.get(end) {
                Some(b'\'' | b'"') => {
                    let quote = bytes[end];
                    end += 1;
                    loop {
                        match bytes.get(end) {
                            Some(b'\\') if bytes.get(end + 1).is_some() => end += 2,
                            Some(byte) if *byte == quote => { end += 1; break; }
                            Some(b'$') if quote == b'"' => return None,
                            Some(_) => end += 1,
                            None => return None,
                        }
                    }
                }
                Some(b'$') => {
                    end += 1;
                    if !bytes.get(end).is_some_and(|b| b.is_ascii_alphabetic() || *b == b'_') {
                        return None;
                    }
                    while bytes.get(end).is_some_and(|b| b.is_ascii_alphanumeric() || *b == b'_') { end += 1; }
                }
                Some(b'0'..=b'9') => {
                    let mut dot = false;
                    while let Some(byte) = bytes.get(end) {
                        if byte.is_ascii_alphanumeric() || *byte == b'_' { end += 1; }
                        else if *byte == b'.' && !dot { dot = true; end += 1; }
                        else { break; }
                    }
                }
                _ => return None,
            }
            while bytes.get(end).is_some_and(u8::is_ascii_whitespace) { end += 1; }
            Some(end)
        };
        let delimiter = |end| matches!(bytes.get(end), None | Some(b';' | b',' | b')' | b']' | b'}' | b'?' | b':'))
            && !(bytes.get(end)==Some(&b':') && bytes.get(end+1)==Some(&b':'))
            && !(bytes.get(end)==Some(&b'?') && bytes.get(end+1..end+3)==Some(b"->"));
        let Some(mut end) = leaf_end(offset) else { return false; };
        let mut saw_operator = false;
        loop {
            if delimiter(end) { return !self.pratt_only || saw_operator; }
            let operator_len = match bytes.get(end) {
                Some(b'<' | b'>') if bytes.get(end + 1) == bytes.get(end) => 2,
                Some(b'<') if bytes.get(end + 1) == Some(&b'=') && bytes.get(end + 2) == Some(&b'>') => return false,
                Some(b'<' | b'>') if bytes.get(end + 1) == Some(&b'=') => 2,
                Some(b'<' | b'>') => 1,
                Some(b'+' | b'-' | b'*' | b'/' | b'%' | b'.') => 1,
                _ => return false,
            };
            saw_operator = true;
            let mut right = end + operator_len;
            while bytes.get(right).is_some_and(u8::is_ascii_whitespace) { right += 1; }
            let Some(next) = leaf_end(right) else { return false; };
            end = next;
        }
        }
    fn finish_rule(&mut self, rule: usize, span: grameo::Span, source: &str) {
        let Some(rule) = Rule::from_capture(grameo::CaptureRule::User(rule)) else {
            return;
        };
        if matches!(rule, Rule::number_lit | Rule::simple_variable | Rule::sq_string | Rule::dq_string) {
            if self
                .arena
                .values()
                .last()
                .is_some_and(|id| matches!(self.arena.get(*id), Value::Marker(_)))
            {
                let raw = &source[span.start..span.end];
                let kind = if rule == Rule::number_lit {
                    numeric_literal::number(&source[span.start..span.end]).kind
                } else if rule == Rule::sq_string {
                    ExprKind::Lit(Literal::Str(decode_php_single_quoted_literal(&raw[1..raw.len() - 1])))
                } else if rule == Rule::dq_string {
                    let body = &raw[1..raw.len() - 1];
                    if body.contains('$') { return; }
                    ExprKind::Lit(Literal::Str(unmask_php_literal_tags_owned(decode_php_double_quoted_literal(body))))
                } else {
                    php_variable_expr_kind(strip_dollar(&source[span.start..span.end]))
                };
                let normalization = leaf_normalization(&kind);
                self.push(Value::Ready {
                    expression: Expression::with_span(kind, self.span(span)),
                    metadata: Metadata {
                        raw: span,
                        anchor: None,
                        pratt_ops: 0,
                        normalization,
                    },
                });
            }
            return;
        }
        if matches!(
            rule,
            Rule::mul_op | Rule::additive_op | Rule::shift_op | Rule::concat_op | Rule::cmp_op
        ) {
            let slot = match parse_binop(&source[span.start..span.end]) {
                BinOp::Add => 0,
                BinOp::Sub => 1,
                BinOp::Mul => 2,
                BinOp::Div => 3,
                BinOp::Mod => 4,
                BinOp::Shl => 5,
                BinOp::Shr => 6,
                BinOp::Concat => 7,
                BinOp::Lt => 8,
                BinOp::Gt => 9,
                BinOp::LtEq => 10,
                BinOp::GtEq => 11,
                BinOp::Spaceship => 12,
                _ => unreachable!("observed arithmetic/comparison operator"),
            };
            self.arena.push(self.operators[slot]);
            return;
        }
        if !selected(rule) {
            return;
        }
        if matches!(rule, Rule::boolean_not_expression | Rule::logic_or_assignment_operand | Rule::logic_and_assignment_operand | Rule::equality_assignment_operand) && !self.pratt_only {
            let values=self.arena.values();
            let frame=values.iter().rposition(|id|matches!(self.arena.get(*id),Value::Marker(marker) if *marker==rule)).expect("active atom frame");
            let ready=values.len()==frame+2 && self.arena.get(values[frame+1]).metadata().is_some_and(|metadata| metadata.raw.start==span.start && metadata.raw.end<=span.end && source[metadata.raw.end..span.end].trim().is_empty());
            if !ready {
                while self.arena.values().len()>frame {self.arena.pop();}
                self.push(Value::Recipe {recipe:ContextRecipe::Atom {rule,raw:span,capture:None},metadata:Metadata {raw:span,anchor:Some(span),pratt_ops:0,normalization:0}});
                return;
            }
        }
        // A certified wrapper changes only the marker below the AST root.
        // Keep that root on the stack and journal one reversible removal.
        let values = self.arena.values();
        if values.len() >= 2 {
            let first = values[values.len() - 1];
            let marker = values[values.len() - 2];
            if matches!(self.arena.get(marker), Value::Marker(marker) if *marker == rule)
                && let Some(metadata) = self.arena.get(first).metadata()
                && metadata.raw.start == span.start
                && metadata.raw.end <= span.end
                && source[metadata.raw.end..span.end].trim().is_empty()
            {
                if metadata.pratt_ops > 0 && metadata.anchor != Some(span) {
                    let ast_span = self.span(span);
                    self.arena.edit(first, |value| {
                        MetadataUndo::ExpressionSpan(std::mem::replace(value.expression_span_mut(), ast_span))
                    }, |value, previous| {
                        let MetadataUndo::ExpressionSpan(span) = previous else { unreachable!() };
                        *value.expression_span_mut() = span;
                    });
                }
                self.set_anchor(first, span);
                self.arena.remove_value_below_top();
                return;
            }
        }
        let Some(first) = self.arena.pop() else {
            return;
        };
        if matches!(self.arena.get(first), Value::Marker(marker) if *marker == rule) {
            return;
        }
        if self.arena.values().last().is_some_and(
            |id| matches!(self.arena.get(*id), Value::Marker(marker) if *marker == rule),
        ) {
            self.arena.pop();
            if let Some(metadata) = self.arena.get(first).metadata()
                && metadata.raw.start == span.start
                && metadata.raw.end <= span.end
                && source[metadata.raw.end..span.end].trim().is_empty()
            {
                if let Some(metadata) = self.arena.get(first).metadata()
                    && metadata.pratt_ops > 0
                    // An existing identical anchor was set after its span edit.
                    // Repeating that edit only adds an undo record; reductions
                    // clear the anchor whenever they construct a new parent.
                    && metadata.anchor != Some(span) {
                    let ast_span = self.span(span);
                    self.arena.edit(first, |value| {
                        MetadataUndo::ExpressionSpan(std::mem::replace(value.expression_span_mut(), ast_span))
                    }, |value, previous| {
                        let MetadataUndo::ExpressionSpan(span) = previous else { unreachable!() };
                        *value.expression_span_mut() = span;
                    });
                }
                self.anchor(first, span);
            }
            return;
        }
        let mut children = vec![first];
        while let Some(id) = self.arena.pop() {
            if matches!(self.arena.get(id), Value::Marker(marker) if *marker == rule) {
                break;
            }
            children.push(id);
        }
        children.reverse();
        if !matches!(
            rule,
            Rule::logic_or_expression | Rule::bit_or_expression | Rule::equality_expression | Rule::logic_or_assignment_operand | Rule::logic_and_assignment_operand | Rule::equality_assignment_operand | Rule::comparison_expression
                | Rule::concat_expression
                | Rule::shift_expression
                | Rule::additive_expression
                | Rule::multiplicative_expression
        ) || children.len() % 2 == 0
        {
            return;
        }
        // Validate before moving any child. Deferred operands and contextual
        // operators keep every surviving AST root and their original captures.
        if !children
            .iter()
            .enumerate()
            .all(|(index, id)| match self.arena.get(*id) {
                Value::Ready { .. } => index % 2 == 0,
                Value::Operator(op) => index % 2 == 1 && owned_operator(*op),
                _ => false,
            })
        {
            return;
        }
        let ast_span = self.span(span);
        let mut left = children[0];
        for pair in children[1..].chunks_exact(2) {
            let Value::Operator(op) = self.arena.get(pair[0]) else {
                unreachable!()
            };
            let op = *op;
            left = self.arena.binary_with_metadata(
                left,
                pair[1],
                |left, right| {
                    let (
                        Value::Ready {
                            expression: left,
                            metadata: left_metadata,
                        },
                        Value::Ready {
                            expression: right,
                            metadata: right_metadata,
                        },
                    ) = (left, right)
                    else {
                        unreachable!()
                    };
                    // The reversible node stores operands without applying folds:
                    // normalization happens after the parse commits, in the live context.
                    let normalization = left_metadata.normalization | right_metadata.normalization
                        | binary_normalization(op, &left, &right);
                    (
                        Value::Ready {
                            expression: Expression::with_span(
                                ExprKind::Binary {
                                    op,
                                    left: Box::new(left),
                                    right: Box::new(right),
                                },
                                ast_span,
                            ),
                            metadata: Metadata {
                                raw: grameo::Span {
                                    start: left_metadata.raw.start,
                                    end: right_metadata.raw.end,
                                },
                                anchor: None,
                                pratt_ops: 0,
                                normalization,
                            },
                        },
                        MetadataUndo::Operands(left_metadata, right_metadata),
                    )
                },
                split_binary,
            );
        }
        self.anchor(left, span);
    }
}

fn split_binary(value: Value, previous: MetadataUndo) -> (Value, Value) {
    if let Value::Recipe {recipe:ContextRecipe::Binary {left,right,..},..}=value {
        let MetadataUndo::RecipeOperands(left_metadata,right_metadata,left_ready,right_ready)=previous else{unreachable!()};
        let restore=|recipe:ContextRecipe,metadata,ready| if ready {let ContextRecipe::Owned(expression)=recipe else{unreachable!()};Value::Ready {expression,metadata}}else{Value::Recipe {recipe,metadata}};
        return (restore(*left,left_metadata,left_ready),restore(*right,right_metadata,right_ready));
    }
    let Value::Ready { expression, .. } = value else {
        unreachable!()
    };
    let ExprKind::Binary { left, right, .. } = expression.kind else {
        unreachable!()
    };
    let MetadataUndo::Operands(left_metadata, right_metadata) = previous else {
        unreachable!()
    };
    (
        Value::Ready {
            expression: *left,
            metadata: left_metadata,
        },
        Value::Ready {
            expression: *right,
            metadata: right_metadata,
        },
    )
}

// The PEG walker gives every reduction in one precedence layer the full
// layer span, including trailing implicit whitespace. Reproduce that layout
// once after a successful Pratt island, before literal folding or binding.
fn recipe_layer(op:BinOp)->u8 {match op {BinOp::BitAnd=>6,BinOp::BitXor=>7,BinOp::BitOr=>8,BinOp::And=>9,BinOp::Or=>10,BinOp::Eq|BinOp::NotEq|BinOp::StrictEq|BinOp::StrictNotEq=>5,BinOp::Concat=>1,BinOp::Shl|BinOp::Shr=>2,BinOp::Add|BinOp::Sub=>3,BinOp::Mul|BinOp::Div|BinOp::Mod=>4,_=>0}}
fn recipe_trivia_end(mut offset:usize,source:&str)->usize {
    let bytes=source.as_bytes();loop {let start=offset;while bytes.get(offset).is_some_and(u8::is_ascii_whitespace){offset+=1;}
        if bytes.get(offset..offset+2)==Some(b"/*") {if let Some(end)=source[offset+2..].find("*/"){offset+=end+4;}else{return offset;}}
        else if bytes.get(offset..offset+2)==Some(b"//") || bytes.get(offset)==Some(&b'#') && bytes.get(offset+1)!=Some(&b'[') {while bytes.get(offset).is_some_and(|b|*b!=b'\n'&&*b!=b'\r'){offset+=1;}}
        if offset==start{return offset;}
    }
}
fn normalize_recipe_spans(recipe:&mut ContextRecipe,source:&str,lines:&[usize]) {
    let mut pending=vec![(recipe,None::<(u8,Span)>)];
    while let Some((recipe,enclosing))=pending.pop(){
        let span_for=|op,mut span:Span| {
            if let Some((parent,inherited))=enclosing && parent==recipe_layer(op) {return inherited;}
            let offset=lines[(span.end_line-1)as usize]+(span.end_col-1)as usize;
            (span.end_line,span.end_col)=offset_to_line_col(recipe_trivia_end(offset,source),lines);span
        };
        match recipe {
            ContextRecipe::Binary {op,span,left,right}=>{*span=span_for(*op,*span);let enclosing=Some((recipe_layer(*op),*span));pending.push((right.as_mut(),enclosing));pending.push((left.as_mut(),enclosing));},
            ContextRecipe::Owned(expression)=>{if let ExprKind::Binary {op,..}=expression.kind {expression.span=span_for(op,expression.span);normalize_pratt_spans(expression,source,lines);}},
            // Contextual captures are complete operand boundaries. Their own
            // walker preserves parentheses and unary/call spans independently.
            ContextRecipe::Atom {..}=>{},
        }
    }
}
fn resolve_recipe(walker:&mut PhpWalker,pair:&Pair<Rule>,recipe:ContextRecipe,module_reads:bool,superglobal_reads:bool)->Result<Expression,String> {
    let input=pair.as_span().get_input();
    let mut atoms:HashMap<(usize,usize,usize),Option<Pair<Rule>>,vybe_runtime::chunk::FxBuildHasher>=HashMap::default();let mut recipes=vec![&recipe];let mut unresolved=false;
    while let Some(recipe)=recipes.pop(){match recipe {ContextRecipe::Atom {rule,raw,capture}=>{let resolved=capture.and_then(|handle|pair.resolve_capture(handle));unresolved|=resolved.is_none();atoms.insert((*rule as usize,raw.start,raw.end),resolved);},ContextRecipe::Binary {left,right,..}=>{recipes.push(left);recipes.push(right);},ContextRecipe::Owned(_)=>{}}}
    if unresolved {
    let mut pending=vec![pair.clone()];
    while let Some(capture)=pending.pop(){let span=capture.as_span();let key=(capture.as_rule() as usize,span.start(),span.end());if let Some(slot)=atoms.get_mut(&key) {*slot=Some(capture);}else{pending.extend(capture.into_inner());}}
    }
    enum Task {Visit(ContextRecipe),Reduce(BinOp,Span)}
    let mut tasks=vec![Task::Visit(recipe)];let mut values=Vec::new();
    while let Some(task)=tasks.pop(){match task {
        Task::Visit(ContextRecipe::Owned(expression))=>values.push(normalize_contextual_binary(expression,walker,module_reads,superglobal_reads)),
        Task::Visit(ContextRecipe::Atom {rule,raw,..})=>{
            let expression=if let Some(Some(capture))=atoms.remove(&(rule as usize,raw.start,raw.end)){
                #[cfg(test)] census_add(5,1);
                walk_expression(walker,capture)?
            }else{
                #[cfg(test)] census_add(6,1);
                let tree=grameo::engine::parse_program_at(&crate::native_grammar::Parser,rule.name(),input,raw.start,grameo::MatchOptions::default()).map_err(|e|e.to_string())?;
                let capture=tree.into_typed_pairs::<Rule>().next().ok_or("missing contextual atom capture")?;
                let previous_input=walker.direct_input.take();let result=walk_expression(walker,capture);walker.direct_input=previous_input;result?
            };values.push(expression);
        },
        Task::Visit(ContextRecipe::Binary {op,span,left,right})=>{tasks.push(Task::Reduce(op,span));tasks.push(Task::Visit(*right));tasks.push(Task::Visit(*left));},
        Task::Reduce(op,span)=>{let right=values.pop().unwrap();let left=values.pop().unwrap();values.push(binary::reduce(walker,left,right,op,"",span));},
    }}
    Ok(values.pop().expect("complete contextual expression recipe"))
}

fn normalize_pratt_spans(expression: &mut Expression, source: &str, lines: &[usize]) {
    fn layer(op: BinOp) -> u8 {
        match op {
            BinOp::BitAnd=>6, BinOp::BitXor=>7, BinOp::BitOr=>8,BinOp::And=>9,BinOp::Or=>10,
            BinOp::Eq|BinOp::NotEq|BinOp::StrictEq|BinOp::StrictNotEq => 5,
            BinOp::Concat => 1,
            BinOp::Shl | BinOp::Shr => 2,
            BinOp::Add | BinOp::Sub => 3,
            BinOp::Mul | BinOp::Div | BinOp::Mod => 4,
            _ => 0,
        }
    }
    let mut pending = vec![(expression, None::<(u8, Span)>)];
    while let Some((expression, enclosing)) = pending.pop() {
        let ExprKind::Binary { op, left, right } = &mut expression.kind else { continue; };
        let precedence = layer(*op);
        let span = if let Some((parent, span)) = enclosing
            && parent == precedence {
            span
        } else {
            let mut span = expression.span;
            let offset = lines[(span.end_line - 1) as usize] + (span.end_col - 1) as usize;
            let whitespace = source.as_bytes()[offset..].iter()
                .take_while(|byte| byte.is_ascii_whitespace()).count();
            (span.end_line, span.end_col) = offset_to_line_col(offset + whitespace, lines);
            span
        };
        expression.span = span;
        pending.push((right.as_mut(), Some((precedence, span))));
        pending.push((left.as_mut(), Some((precedence, span))));
    }
}

#[inline]
pub(super) fn take_for_pair(walker: &mut PhpWalker, pair: &Pair<Rule>) -> Result<Option<Expression>, String> {
    let span = pair.as_span();
    let input = span.get_input();
    // Embedded expressions can be reparsed from a substring during contextual
    // lowering. Their offsets belong to that input, not the module's captures.
    if walker.direct_input != Some((input.as_ptr() as usize, input.len())) {
        return Ok(None);
    }
    // Argument inference can inspect a capture again before generic dispatch.
    // Reuse the parsed value and bind it in the live scope on every inspection.
    let superglobal_reads = walker.assign_lhs_depth == 0;
    let module_reads = walker.function_stack.is_empty() && superglobal_reads;
    let key = (span.start(), span.end());
    if walker.direct_expressions.get(&key).is_some_and(|entry|matches!(entry,CachedExpression::Contextual(Some(_)))) {
        let mut cached=walker.direct_expressions.remove(&key).unwrap();
        let CachedExpression::Contextual(recipe)=&mut cached else{unreachable!()};
        let result=resolve_recipe(walker,pair,recipe.take().unwrap(),module_reads,superglobal_reads);
        walker.direct_expressions.insert(key,cached);
        return result.map(|expression| Some(php_bind_symbol_scope_expression(expression, walker)));
    }
    let contextual = walker.direct_expressions.get(&key).is_some_and(|expression| {
        matches!(expression, CachedExpression::Deferred { first: Some(_), normalization }
            if *normalization & NEEDS_CONTEXTUAL_BINARY != 0)
    });
    if contextual {
        let mut cached = walker.direct_expressions.remove(&key).unwrap();
        let CachedExpression::Deferred { first, .. } = &mut cached else { unreachable!() };
        let expression = normalize_contextual_binary(first.take().unwrap(), walker, module_reads, superglobal_reads);
        walker.direct_expressions.insert(key, cached);
        return Ok(Some(php_bind_symbol_scope_expression(expression, walker)));
    }
    let Some(expression) = walker.direct_expressions.get_mut(&key) else {
        return Ok(None);
    };
    if !matches!(expression, CachedExpression::Deferred { first: None, .. } | CachedExpression::Contextual(None)) {
        let expression = expression.bind(input, module_reads, superglobal_reads, &walker.line_starts);
        return Ok(Some(php_bind_symbol_scope_expression(expression, walker)));
    }
    // Reconstruct only on an actual repeat. Full input keeps absolute spans,
    // and the same walker retains the caller's scope and directive policy.
    let tree = grameo::engine::parse_program_at(
        &crate::native_grammar::Parser, pair.as_rule().name(), input,
        span.start(), grameo::MatchOptions::default(),
    ).map_err(|error| error.to_string())?;
    let recaptured = tree.into_typed_pairs::<Rule>().next()
        .ok_or_else(|| "owned expression recapture produced no capture".to_owned())?;
    if recaptured.as_span().start() != span.start() || recaptured.as_span().end() != span.end() {
        return Err("owned expression recapture changed its source span".to_owned());
    }
    let previous_input = walker.direct_input.take();
    let result = walk_expression(walker, recaptured);
    walker.direct_input = previous_input;
    result.map(Some)
}

// Bind scope-dependent reductions once, using the capture walker's existing
// semantic hook. Parser construction remains transactional and owns operands.
fn normalize_contextual_binary(expression: Expression, walker: &mut PhpWalker, module_reads: bool, superglobal_reads: bool) -> Expression {
    enum Task { Visit(Expression), Reduce(BinOp, Span, Box<Expression>, Box<Expression>) }
    let mut tasks = vec![Task::Visit(expression)];
    let mut values = Vec::new();
    while let Some(task) = tasks.pop() {
        match task {
            Task::Visit(Expression { kind: ExprKind::Binary { op, mut left, mut right }, span }) => {
                let placeholder = || Expression::with_span(ExprKind::Lit(Literal::Null), span);
                let left_value = std::mem::replace(left.as_mut(), placeholder());
                let right_value = std::mem::replace(right.as_mut(), placeholder());
                tasks.push(Task::Reduce(op, span, left, right));
                tasks.push(Task::Visit(right_value));
                tasks.push(Task::Visit(left_value));
            }
            Task::Visit(expression) => values.push(bind_owned_leaf(expression, module_reads, superglobal_reads)),
            Task::Reduce(op, span, mut left, mut right) => {
                *right = values.pop().unwrap();
                *left = values.pop().unwrap();
                values.push(if matches!(op, BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq | BinOp::Eq | BinOp::NotEq | BinOp::StrictEq | BinOp::StrictNotEq | BinOp::BitAnd | BinOp::BitXor | BinOp::BitOr | BinOp::And | BinOp::Or) {
                    binary::reduce(walker, *left, *right, op, "", span)
                } else {
                    binary::normalize_arithmetic(Expression::with_span(ExprKind::Binary { op, left, right }, span))
                });
            }
        }
    }
    values.pop().unwrap()
}

// Normalize committed arithmetic with ownership of the original operand boxes.
// The explicit stack also keeps long expression chains off the native stack.
#[cfg(test)]
fn normalize_owned(expression: Expression) -> (Expression, bool) {
    transform_owned(expression, None)
}
#[cfg(test)]
fn transform_owned(expression: Expression, module_reads: Option<bool>) -> (Expression, bool) {
    let (value, contextual, _) = transform_owned_impl(expression, module_reads, false, None);
    (value, contextual)
}
fn transform_owned_impl(
    expression: Expression,
    module_reads: Option<bool>,
    superglobal_reads: bool,
    replay_source: Option<(&str, &[usize])>,
) -> (Expression, bool, Vec<Replay>) {
    let mut replay = Vec::new();
    if !matches!(expression.kind, ExprKind::Binary { .. }) {
        if let Some((source, lines)) = replay_source {
            replay.push(record_replay_leaf(&expression, source, lines));
        }
        return (bind_owned_leaf(expression, module_reads.unwrap_or(false), superglobal_reads), false, replay);
    }
    enum Task {
        Visit(Expression),
        Reduce(BinOp, Span, Box<Expression>, Box<Expression>),
    }
    let mut tasks = vec![Task::Visit(expression)];
    let mut values = Vec::new();
    while let Some(task) = tasks.pop() {
        match task {
            Task::Visit(Expression { kind: ExprKind::Binary { op, mut left, mut right }, span }) => {
                let placeholder = || Expression::with_span(ExprKind::Lit(Literal::Null), span);
                let left_value = std::mem::replace(left.as_mut(), placeholder());
                let right_value = std::mem::replace(right.as_mut(), placeholder());
                tasks.push(Task::Reduce(op, span, left, right));
                tasks.push(Task::Visit(right_value));
                tasks.push(Task::Visit(left_value));
            }
            Task::Visit(expression) => {
                if let Some((source, lines)) = replay_source {
                    replay.push(record_replay_leaf(&expression, source, lines));
                }
                let contextual = module_reads.is_none()
                    && matches!(&expression.kind, ExprKind::Ident(name) if php_request_global_key(name).is_some());
                values.push((bind_owned_leaf(expression, module_reads.unwrap_or(false), superglobal_reads), contextual));
            }
            Task::Reduce(op, span, mut left, mut right) => {
                if replay_source.is_some() { replay.push(Replay::Binary { op, span }); }
                let (right_value, right_contextual) = values.pop().unwrap();
                let (left_value, left_contextual) = values.pop().unwrap();
                *right = right_value;
                *left = left_value;
                let contextual = left_contextual || right_contextual;
                let expression = Expression::with_span(ExprKind::Binary { op, left, right }, span);
                // Variable names can resemble PHP_INT_MAX/MIN. Apply folds
                // after scope binding whenever an operand requires context.
                values.push((if contextual { expression } else {
                    binary::normalize_arithmetic(expression)
                }, contextual));
            }
        }
    }
    let (value, contextual) = values.pop().unwrap();
    (value, contextual, replay)
}

fn bind_owned_leaf(expression: Expression, module_reads: bool, superglobal_reads: bool) -> Expression {
    if superglobal_reads
        && let ExprKind::Ident(name) = &expression.kind
        && PHP_SUPERGLOBAL_NAMES.contains(&name.as_str()) {
        let span = expression.span;
        return Expression::with_span(ExprKind::RefLoad(Box::new(expression)), span);
    }
    if module_reads
        && let ExprKind::Ident(name) = &expression.kind
        && let Some(key) = php_request_global_key(name) {
        return php_module_global_read_expr(name, key, &expression.span);
    }
    expression
}

fn record_replay_leaf(expression: &Expression, source: &str, lines: &[usize]) -> Replay {
    let from_source = || -> Option<Replay> {
        let span = expression.span;
        let start = *lines.get(span.start_line.checked_sub(1)? as usize)?
            + span.start_col.checked_sub(1)? as usize;
        let end = *lines.get(span.end_line.checked_sub(1)? as usize)?
            + span.end_col.checked_sub(1)? as usize;
        let text = source.get(start..end)?;
        let valid = match &expression.kind {
            ExprKind::Ident(name) => text.starts_with('$') && name == strip_dollar(text),
            ExprKind::Lit(Literal::Str(_)) => text.starts_with('\'') || text.starts_with('"'),
            ExprKind::Lit(_) => text.as_bytes().first().is_some_and(u8::is_ascii_digit),
            _ => false,
        };
        valid.then_some(Replay::Leaf { raw: grameo::Span { start, end }, span })
    };
    from_source().unwrap_or_else(|| Replay::Value(Box::new(expression.clone())))
}

fn replay_expression(replay: &[Replay], source: &str, module_reads: bool, superglobal_reads: bool) -> Expression {
    let mut values = Vec::new();
    for instruction in replay {
        match instruction {
            Replay::Value(expression) => values.push(bind_ready_leaf(expression, module_reads, superglobal_reads)),
            Replay::Leaf { raw, span } => {
                let text = &source[raw.start..raw.end];
                let kind = match text.as_bytes()[0] {
                    b'$' => ExprKind::Ident(strip_dollar(text).to_owned()),
                    b'\'' => ExprKind::Lit(Literal::Str(decode_php_single_quoted_literal(&text[1..text.len() - 1]))),
                    b'"' => ExprKind::Lit(Literal::Str(unmask_php_literal_tags_owned(decode_php_double_quoted_literal(&text[1..text.len() - 1])))),
                    _ => numeric_literal::number(text).kind,
                };
                values.push(bind_owned_leaf(Expression::with_span(kind, *span), module_reads, superglobal_reads));
            }
            Replay::Binary { op, span } => {
                let right = Box::new(values.pop().unwrap());
                let left = Box::new(values.pop().unwrap());
                values.push(binary::normalize_arithmetic(Expression::with_span(
                    ExprKind::Binary { op: *op, left, right }, *span)));
            }
        }
    }
    values.pop().unwrap()
}

// Copy the reusable parsed expression and apply contextual binding in one walk.
// Cached nodes contain only arithmetic expressions and context-free leaves.
#[cfg(test)]
fn bind_ready_expression(expression: &Expression, module_reads: bool, normalize: bool) -> Expression {
    let superglobal_reads = false;
    if !matches!(expression.kind, ExprKind::Binary { .. }) {
        return bind_ready_leaf(expression, module_reads, superglobal_reads);
    }
    enum Task<'a> {
        Visit(&'a Expression),
        Reduce(BinOp, Span),
    }
    let mut tasks = vec![Task::Visit(expression)];
    let mut values = Vec::new();
    while let Some(task) = tasks.pop() {
        match task {
            Task::Visit(Expression {
                kind: ExprKind::Binary { op, left, right },
                span,
            }) => {
                tasks.push(Task::Reduce(*op, *span));
                tasks.push(Task::Visit(right));
                tasks.push(Task::Visit(left));
            }
            Task::Visit(expression) => {
                values.push(bind_ready_leaf(expression, module_reads, superglobal_reads));
            }
            Task::Reduce(op, span) => {
                let right = Box::new(values.pop().unwrap());
                let left = Box::new(values.pop().unwrap());
                let expression = Expression::with_span(
                    ExprKind::Binary { op, left, right },
                    span,
                );
                values.push(if normalize { binary::normalize_arithmetic(expression) } else { expression });
            }
        }
    }
    values.pop().unwrap()
}

fn bind_ready_leaf(expression: &Expression, module_reads: bool, superglobal_reads: bool) -> Expression {
    bind_owned_leaf(expression.clone(), module_reads, superglobal_reads)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unbound_source_intents_keep_plain_ast_capture_behavior() {
        let source = "return 1 + (include 'a');";
        let mut ordinary = Ast::new(source);
        let mut unbound = Ast::new(source);
        let mark = unbound.checkpoint();
        unbound.enable_source_intents(Vec::new());
        assert!(unbound.source_hooks.is_none());
        assert_eq!(unbound.checkpoint(), mark);
        let plain = crate::native_grammar::Parser.build_captures(Rule::program_pure, source, &mut ordinary).unwrap();
        let actual = crate::native_grammar::Parser.build_captures(Rule::program_pure, source, &mut unbound).unwrap();
        assert_eq!(format!("{plain:?}"), format!("{actual:?}"));
        assert!(unbound.take_source_intents().is_empty());
    }

    #[test]
    fn certified_wrapper_restores_stack_anchor_and_span_on_rollback() {
        let position = |span: Span| (span.start_line, span.start_col, span.end_line, span.end_col);
        for pratt_ops in [0, 1] {
            let source = "1 ";
            let raw = grameo::Span { start: 0, end: 1 };
            let wrapper = grameo::Span { start: 0, end: 2 };
            let mut ast = Ast::new(source);
            ast.push(Value::Marker(Rule::additive_expression));
            let expression = Expression::with_span(numeric_literal::number("1").kind, ast.span(raw));
            let original_span = expression.span;
            ast.push(Value::Ready { expression, metadata: Metadata {
                raw, anchor: None, pratt_ops, normalization: 0,
            } });
            let original_stack = ast.arena.values().to_vec();
            let root = original_stack[1];
            let mark = ast.checkpoint();
            for _ in 0..3 {
                ast.finish_rule(Rule::additive_expression as usize, wrapper, source);
                assert_eq!(ast.arena.values(), &[root]);
                let Value::Ready { expression, metadata } = ast.arena.get(root) else { panic!("AST root lost") };
                assert_eq!(metadata.anchor, Some(wrapper));
                assert_eq!(position(expression.span), position(if pratt_ops == 0 { original_span } else { ast.span(wrapper) }));
                ast.rollback(mark);
                assert_eq!(ast.arena.values(), original_stack);
                let Value::Ready { expression, metadata } = ast.arena.get(root) else { panic!("AST root lost") };
                assert_eq!(metadata.anchor, None);
                assert_eq!(position(expression.span), position(original_span));
            }
        }
    }

    #[test]
    fn deferred_compound_recapture_preserves_live_scope_and_spans() {
        for source in ["$request + 1", "$a + $b * 3", "'hello' . 'world'", "1 + 2 + 3",
            "$_GET + 1", "1 / 0", "9223372036854775807 + 1", "-9223372036854775807 - 2",
            "(9223372036854775807 + 1) * 2", "(1 / 0) + 2"] {
            let mut walker = PhpWalker::default();
            let pair = parse_pairs_with_frontend(&mut walker, true, false,
                Rule::comparison_expression, source).unwrap().next().unwrap();
            for lhs_depth in [0, 1, 0] {
                walker.assign_lhs_depth = lhs_depth;
                let actual = walk_expression(&mut walker, pair.clone()).unwrap();
                let mut reference = PhpWalker::default();
                reference.assign_lhs_depth = lhs_depth;
                let peg = PhpParser::parse(Rule::comparison_expression, source).unwrap().next().unwrap();
                let expected = walk_expression(&mut reference, peg).unwrap();
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
                assert!(walker.direct_input.is_some());
            }
        }
    }

    #[test]
    fn normalization_facts_are_composed_during_owned_construction() {
        for (source, expected) in [
            ("$a * 3 + 1", NEEDS_GLOBAL_BINDING),
            ("3 * 2 + 1", 0),
            ("9223372036854775807 + 1", NEEDS_ARITHMETIC_FOLD),
            ("1 / 0", NEEDS_ARITHMETIC_FOLD),
            ("$_GET + 1", NEEDS_GLOBAL_BINDING | NEEDS_SUPERGLOBAL_BINDING),
        ] {
            let mut walker = PhpWalker::default();
            let pair = parse_pairs_with_frontend(&mut walker, true, false,
                Rule::comparison_expression, source).unwrap().next().unwrap();
            let span = pair.as_span();
            let cached = walker.direct_expressions.get(&(span.start(), span.end())).unwrap();
            let CachedExpression::Deferred { normalization, .. } = cached else {
                panic!("expected an owned compound: {source}");
            };
            assert_eq!(*normalization, expected, "{source}");
        }
    }

    #[test]
    fn mode_and_ast_journals_restore_each_interleaved_boundary() {
        let mut builder = Ast::pratt_only("$a + 1;");
        let mut states = Vec::new();
        for i in 0..128 {
            states.push((builder.checkpoint(), builder.modes.clone(), builder.arena.values().to_vec()));
            match i % 5 {
                0 => builder.push_mode(i % 2 == 0),
                1 => {
                    let id = builder.arena.alloc(Value::Operator(BinOp::Add));
                    builder.arena.push(id);
                }
                2 => { builder.arena.pop(); }
                3 => builder.pop_mode(),
                _ => builder.push_mode(false),
            }
        }
        for (mark, modes, values) in states.into_iter().rev() {
            builder.rollback(mark);
            assert_eq!(builder.modes, modes);
            assert_eq!(builder.arena.values(), values);
            builder.rollback(mark);
            assert_eq!(builder.checkpoint(), mark);
            assert_eq!(builder.modes, modes);
        }
        assert!(!builder.mode_enabled());
        assert!(builder.mode_undo.is_empty());
    }

    #[test]
    fn pratt_only_modes_are_transactional_and_preserve_contextual_fallback() {
        for source in [
            "<?php $x = $a + 1; echo $x;",
            "<?php $x = f($a + 1) + 2; $y = $o->f() * 3;",
            "<?php $x = (1 + 2) * 3; $y = -2 ** 3 + 4;",
            "<?php function f($a) { return $a + 1; } $a = 2; echo f($a + 2);",
            "<?php $x = $a < $b; $y = $a + 1 /* comment */ + 2;",
            "<?php $x = $GLOBALS + 1; $y = $this * 2; echo \"interpolated $x\";",
        ] {
            let expected = super::super::parse_with_frontend(source, false).unwrap();
            let actual = super::super::parse_direct_pratt(source).unwrap();
            assert_modules(&actual, &expected, source);
        }
        for source in ["$a + ;", "f($a + 1", "$a * f("] {
            let mut builder = Ast::pratt_only(source);
            let mark = builder.checkpoint();
            assert!(grameo::engine::build_capture_program(
                &crate::native_grammar::Parser, "program_pure", source,
                grameo::MatchOptions::default(), &mut builder,
            ).is_err());
            assert_eq!(builder.checkpoint(), mark);
            assert!(!builder.mode_enabled());
            builder.rollback(mark);
            assert!(!builder.mode_enabled());
        }
    }

    #[test]
    fn fused_replay_preserves_unsupported_leaves_and_live_scope_binding() {
        let source = "$x + 1";
        let lines = build_line_starts(source);
        let variable = Expression::with_span(ExprKind::Ident("x".into()), Span {
            start_line: 1, start_col: 1, end_line: 1, end_col: 3,
        });
        // An original AST leaf that cannot be reconstructed from source text.
        // Retain that leaf in the recipe instead of a second full AST.
        let value = Expression::new(ExprKind::Lit(Literal::Int(7)));
        let expression = Expression::new(ExprKind::Binary {
            op: BinOp::Add, left: Box::new(variable), right: Box::new(value),
        });
        let mut cache = CachedExpression::Pending(Some(expression.clone()));
        for module_reads in [true, false, true, false] {
            let expected = bind_ready_expression(&expression, module_reads, true);
            assert_eq!(format!("{:?}", cache.bind(source, module_reads, false, &lines)), format!("{expected:?}"));
        }
        let CachedExpression::Owned(replay) = cache else { panic!("owned replay") };
        assert!(replay.iter().any(|item| matches!(item, Replay::Value(_))));
    }
    #[test]
    fn owned_binary_cache_transfers_boxes_and_replays_in_live_scope() {
        let source = "$a + $b";
        let builder = Ast::new(source);
        let left = Box::new(Expression::with_span(php_variable_expr_kind("$a"), builder.span(grameo::Span { start: 0, end: 2 })));
        let right = Box::new(Expression::with_span(php_variable_expr_kind("$b"), builder.span(grameo::Span { start: 5, end: 7 })));
        let left_address = left.as_ref() as *const Expression;
        let right_address = right.as_ref() as *const Expression;
        let raw = grameo::Span { start: 0, end: source.len() };
        let expression = Expression::with_span(ExprKind::Binary { op: BinOp::Add, left, right }, builder.span(raw));
        let mut cache = CachedExpression::new(expression, raw, source, &builder.lines);
        assert!(matches!(cache, CachedExpression::Pending(Some(_))));
        let first = cache.bind(source, false, false, &builder.lines);
        let ExprKind::Binary { left, right, .. } = &first.kind else { panic!("binary expected") };
        assert_eq!(left.as_ref() as *const Expression, left_address);
        assert_eq!(right.as_ref() as *const Expression, right_address);
        assert_eq!(format!("{first:?}"), format!("{:?}", cache.bind(source, false, false, &builder.lines)));

        let source = "$request + 1";
        let builder = Ast::new(source);
        let raw = grameo::Span { start: 0, end: source.len() };
        let expression = Expression::with_span(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::with_span(php_variable_expr_kind("$request"), builder.span(grameo::Span { start: 0, end: 8 }))),
            right: Box::new(Expression::with_span(ExprKind::Lit(Literal::Int(1)), builder.span(grameo::Span { start: 11, end: 12 }))),
        }, builder.span(raw));
        let expected_local = bind_ready_expression(&expression, false, true);
        let expected_module = bind_ready_expression(&expression, true, true);
        assert_ne!(format!("{expected_local:?}"), format!("{expected_module:?}"));
        let mut cache = CachedExpression::new(expression, raw, source, &builder.lines);
        assert_eq!(format!("{:?}", cache.bind(source, true, false, &builder.lines)), format!("{expected_module:?}"));
        assert_eq!(format!("{:?}", cache.bind(source, false, false, &builder.lines)), format!("{expected_local:?}"));
        assert_eq!(format!("{:?}", cache.bind(source, true, false, &builder.lines)), format!("{expected_module:?}"));
    }
    #[test]
    fn literal_normalization_preserves_escapes_unicode_and_owned_storage() {
        for (body, expected) in [
            ("plain 雪", "plain 雪"),
            (r"a\'b", "a'b"),
            (r"a\\b", r"a\b"),
            (r"a\nb", r"a\nb"),
            (r"\n", "\n"),
            (r"\é\q", r"\é\q"),
            ("\u{E000}\u{E001}php\u{E001}\u{E000}", "<?php?>"),
        ] {
            assert_eq!(decode_php_single_quoted_literal(body), expected);
        }
        let value = "ordinary SQL text 雪".to_owned();
        let address = value.as_ptr();
        let unmasked = unmask_php_literal_tags_owned(value);
        assert_eq!(unmasked.as_ptr(), address);
    }
    #[test]
    fn pratt_reduction_cannot_cross_an_unsupported_operand() {
        let source = "$a * missing + 1";
        let mut builder = Ast::new(source);
        for (kind, raw) in [
            (ExprKind::Ident("$a".into()), grameo::Span { start: 0, end: 2 }),
            (ExprKind::Lit(Literal::Int(1)), grameo::Span { start: source.len() - 1, end: source.len() }),
        ] {
            builder.push(Value::Ready {
                expression: Expression::with_span(kind, builder.span(raw)),
                metadata: Metadata { raw, anchor: None, pratt_ops: 0, normalization: 0 },
            });
        }
        let before = builder.arena.values().to_vec();
        let start = source.find('+').unwrap();
        builder.reduce(Rule::additive_op as usize, grameo::pratt::Fixity::InfixLeft,
            grameo::Span { start, end: start + 1 });
        assert_eq!(builder.arena.values(), before);
        assert!(!builder.accepts_pratt(Rule::comparison_expression as usize,
            grameo::Span { start: 0, end: source.len() }));
    }
    #[test]
    fn borrowed_source_normalization_preserves_exact_parser_input() {
        for source in ["<?php echo 1;\n", "<?php echo 1;", "<?php // eof comment", "<?php echo '?>';\n", "html <?php echo 1; ?> tail"] {
            assert_eq!(normalized_source_for_parse(source).as_ref(), normalize_mixed_php_source(source));
        }
        assert!(matches!(normalized_source_for_parse("<?php echo 1;\n"), std::borrow::Cow::Borrowed(_)));
    }
    #[test]
    fn plain_text_before_php_tags_uses_the_same_normalized_ast() {
        let source = "html<?php $x=1;?>é<b><?= $x+2?></b>";
        let normalized = normalize_mixed_php_source(source);
        for parser in [parse, parse_direct] {
            let actual = parser(source).expect("mixed HTML after plain text");
            let expected = parser(&normalized).expect("normalized mixed HTML");
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
            assert!(parser("echo '<?';").is_ok(), "a tag inside valid PHP stays a literal");
        }
    }
    #[test]
    fn committed_normalization_keeps_operand_storage_and_contextual_reads() {
        let span = Span::default();
        let left = Box::new(Expression::with_span(ExprKind::Ident("amount".into()), span));
        let right = Box::new(Expression::with_span(ExprKind::Lit(Literal::Int(1)), span));
        let left_address = left.as_ref() as *const Expression;
        let right_address = right.as_ref() as *const Expression;
        let (expression, contextual) = normalize_owned(Expression::with_span(
            ExprKind::Binary { op: BinOp::Add, left, right }, span,
        ));
        assert!(contextual);
        let ExprKind::Binary { left, right, .. } = expression.kind else { panic!("expected arithmetic"); };
        assert_eq!(left.as_ref() as *const Expression, left_address);
        assert_eq!(right.as_ref() as *const Expression, right_address);
    }
    fn assert_modules(actual: &Module, expected: &Module, label: &str) {
        let actual = format!("{actual:?}");
        let expected = format!("{expected:?}");
        if actual != expected {
            let offset = actual
                .chars()
                .zip(expected.chars())
                .position(|(a, b)| a != b)
                .unwrap_or(actual.chars().count().min(expected.chars().count()));
            let snippet = |value: &str| {
                value
                    .chars()
                    .skip(offset.saturating_sub(80))
                    .take(240)
                    .collect::<String>()
            };
            panic!(
                "{label}: AST differs at character {offset}; actual={} expected={}",
                snippet(&actual),
                snippet(&expected)
            );
        }
    }
    #[test]
    fn php_direct_ast_preserves_complete_modules_and_source_spans() {
        for source in [
            "<?php $a = 1 + 2 * 3; echo $a - 4 / 2;",
            "<?php $a = $x + $y * $z - 4 / $w; $b = $x << 2 + $y; echo 'prefix' . $x . 'suffix' . $y;",
            "<?php $a = $x + 1 + 2 + 3 + 4; $b = $x + unknown() + $y; $c = $x * self::N + $y;",
            "<?php $x = $a + $b ; $y = 1 / 0 ; $z = $a << 2; echo $a . $b;",
            "<?php $x = $a + foo(); $y = $a + SOME_CONST; $z = $a + $b * $c; $w = $a < $b;",
            r#"<?php $x = 'SQL 雪 \\path \'quote'; echo 'a' . 'b'; $date = new DateTime('now'); echo $date->format('Y-m-d'); echo "plain\ntext"; echo "interpolated $x";"#,
            "<?php $a = (1 + 2) * 3; $b = -2 ** 3 + 4;",
            "<?php $a = 9223372036854775807 + 1; $b = 1 / 0;",
            "<?php $PHP_INT_MAX = 4; $a = $PHP_INT_MAX + 1; function f($PHP_INT_MAX) { return $PHP_INT_MAX + 1; }",
            "<?php $a = f($x + 1) + 2; $b = $obj->f()[0] * 3;",
            "<?php namespace Demo; class A { const N = 2 + 3; function f($x = 1+2) { return $x * self::N + 1; } }",
            "<?php $a = $GLOBALS + 1; $b = $this * 2; $c = $a < $b;",
            "<?php $a = 0xF + 0b11 * 0o7; $b = 1_000 . 2 + 3;",
            "<?php $a = 1 /* comment */ + 2; $b = [$a + 3, f(4 + 5)];",
            "hello <?php $x = 1 + 2; echo $x; ?> world",
            "<?php function f($format, $date) { return (new DateTime($date, new DateTimeZone('UTC')))->format($format); }",
            "<?php $date = new DateTime('now'); echo $date->format(123); echo $date->format(0x10);",
            "<?php $x = 2; function f($x) { global $g; return $x + $g; } echo \"value={$x}\";",
        ] {
            let expected = super::super::parse_with_frontend(source, false).unwrap();
            let actual = super::super::parse_with_frontend(source, true).unwrap();
            assert_modules(&actual, &expected, source);
        }
    }
    #[test]
    fn relational_pratt_keeps_contextual_normalization_and_precedence() {
        for op in ["<", ">", "<=", ">="] {
            for operands in [("$a + 1", "$b * 2"), ("1", "2"), ("'2'", "'10'"), ("$_GET", "$a"), ("$date", "$other")] {
                for context in ["$out = EXPR;", "echo EXPR;", "if (EXPR) { echo 1; }", "function f($a, $b) { global $date, $other; return EXPR; }"] {
                    let expression = format!("{} {op} {}", operands.0, operands.1);
                    let source = format!("<?php $a = 1; $b = 2; $date = new DateTime('2020-01-01'); $other = new DateTime('2021-01-01'); {}", context.replace("EXPR", &expression));
                    let expected = super::super::parse(&source).unwrap();
                    let actual = super::super::parse_direct(&source).unwrap();
                    assert_modules(&actual, &expected, &source);
                }
            }
        }
        for source in ["<?php $x = $a < $b < $c;", "<?php $x = 1 <=> 2;", "<?php $x = $a /* comment */ <= $b;", "<?php $x = ($a < $b) + 1;", "<?php $x = $a <=;", "<?php $x = $a <=>;"] {
            assert_eq!(super::super::parse_direct(source).map(|m| format!("{m:?}")), super::super::parse(source).map(|m| format!("{m:?}")), "{source}");
        }
    }
    #[test]
    fn php_direct_ast_preserves_real_wordpress_modules() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/webroot/wordpress/wp-includes");
        for name in [
            "class-wp-hook.php",
            "class-wp-query.php",
            "formatting.php",
            "functions.php",
        ] {
            let source = std::fs::read_to_string(root.join(name)).unwrap();
            let expected = super::super::parse_with_frontend(&source, false).unwrap();
            let actual = super::super::parse_with_frontend(&source, true).unwrap();
            assert_modules(&actual, &expected, name);
        }
    }
    #[test]
    fn php_direct_ast_failures_restore_owned_values_and_metadata() {
        for source in ["$x = 1 + 2 *;", "$x = f(1 + 2", "$x = 1 + ;"] {
            let mut builder = Ast::new(source);
            let before = builder.checkpoint();
            assert!(
                grameo::engine::build_capture_program(
                    &crate::native_grammar::Parser,
                    "program_pure",
                    source,
                    grameo::MatchOptions::default(),
                    &mut builder
                )
                .is_err()
            );
            assert_eq!(builder.checkpoint(), before);
            assert!(builder.finish().0.is_empty());
        }
    }
    #[test]
    fn selective_comparison_limits_restore_all_mode_and_value_state() {
        for source in ["$a[0] + f($b + 2);", "f($a + 1, $o->method($b * 2));", "$a + f($b + ;", "$a + 1;"] {
            for depth in 0..24 {
                for work in 0..256 {
                    let mut builder = Ast::new(source);
                    let mark = builder.checkpoint();
                    let result = grameo::engine::build_capture_program(
                        &crate::native_grammar::Parser, "program_pure", source,
                        grameo::MatchOptions { max_rule_depth: depth, max_steps: work, ..Default::default() }, &mut builder,
                    );
                    if result.is_err() {
                        assert_eq!(builder.checkpoint(), mark, "{source} depth={depth} work={work}");
                        assert!(builder.modes.is_empty());
                        assert!(builder.mode_undo.is_empty());
                        assert!(builder.finish().0.is_empty());
                    }
                }
            }
        }
    }

    #[test]
    fn flattened_equality_bitwise_preserves_operator_trivia_and_assignment_precedence() {
        for source in [
            "<?php return (6 & 3) | 8;",
            "<?php return 7 ^ 3 & 1;",
            "<?php return 1 | 2 ^ 3 & 1;",
            "<?php $x=0; return false !== $x=4 | 1;",
            "<?php return (2 === 2) & 1;",
            "<?php return ord('AB' & 'BC');",
            "<?php return (6 & /* guard trivia */ 3) | /* guard trivia */ 8;",
            "<?php return 7 ^ /* gap */ 3 & /* guard trivia */ 1;",
            "<?php return 6 & /* missing operand */ ;",
            "<?php return 6 | /* unterminated",
        ] {
            assert_eq!(super::super::parse_direct(source).map(|module|format!("{module:?}")),super::super::parse(source).map(|module|format!("{module:?}")),"{source}");
        }
    }

    #[test]
    fn contextual_pratt_atoms_keep_calls_scope_comments_and_parenthesis_boundaries() {
        for source in [
            "<?php function f($a){return $a+1;} $a=[2,3]; $x=f($a[0])+($a[1]*2)+ /* gap */ -3; return $x;",
            "<?php class Box { public $x=2; function f($a){return $this->x+$a;} } $o=new Box; return $o->f(3)+$o->x;",
            "<?php $x=4; $f=function($y)use($x){return $x+$y;}; return $f(3)+($x-1);",
            "<?php function f(){global $a; return $a+1;} $a=2; $b=f()+$GLOBALS['x'];",
            "<?php $a=[1,2]; $b=($a[0]+2)+($a[1]+3); $c=(($a[0]+2)+3)+f();",
            "<?php $a=f() /* trailing atom */ + /* operator gap */ g() * (h()+2);",
            "<?php $x=$a <=> $b; $y=f() <=> g(); $z=f() == g();",
            "<?php class Foo {static function f($x){return $x;}} $class='Foo'; return $class::f(1)+2;",
            "<?php $o=null; return $o?->x ?? 3;",
            "<?php $x=f()+;", "<?php $x=f(1+2;", "<?php $x=($a+$b)+;",
        ] {
            assert_eq!(super::super::parse_direct(source).map(|module|format!("{module:?}")),super::super::parse(source).map(|module|format!("{module:?}")),"{source}");
        }
        let source="$x = f($a+1) + $o->field[0];";
        let mut builder=Ast::new(source);
        grameo::engine::build_capture_program(&crate::native_grammar::Parser,"program_pure",source,grameo::MatchOptions::default(),&mut builder).unwrap();
        assert!(builder.finish().0.values().any(|value|matches!(value,CachedExpression::Contextual(Some(ContextRecipe::Binary {..})))));
    }
    #[test]
    fn mixed_contextual_recipe_reduction_restores_exact_operand_representations() {
        let source="$a + f()";let mut builder=Ast::new(source);
        builder.push(Value::Ready { expression:Expression::with_span(ExprKind::Ident("a".into()),builder.span(grameo::Span {start:0,end:2})),metadata:Metadata {raw:grameo::Span {start:0,end:2},anchor:None,pratt_ops:0,normalization:NEEDS_GLOBAL_BINDING}});
        builder.push(Value::Recipe {recipe:ContextRecipe::Atom {rule:Rule::boolean_not_expression,raw:grameo::Span {start:5,end:8},capture:None},metadata:Metadata {raw:grameo::Span {start:5,end:8},anchor:None,pratt_ops:0,normalization:0}});
        let original=builder.arena.values().to_vec();let checkpoint=builder.checkpoint();
        let Value::Ready {expression:Expression {kind:ExprKind::Ident(name),..},..}=builder.arena.get(original[0]) else{unreachable!()};let pointer=name.as_ptr();
        builder.reduce(Rule::additive_op as usize,grameo::pratt::Fixity::InfixLeft,grameo::Span {start:3,end:4});
        assert!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()),Value::Recipe {recipe:ContextRecipe::Binary {..},..}));
        builder.rollback(checkpoint);assert_eq!(builder.arena.values(),original);
        let Value::Ready {expression:Expression {kind:ExprKind::Ident(name),..},metadata}=builder.arena.get(original[0]) else{unreachable!()};assert_eq!(name.as_ptr(),pointer);assert_eq!(metadata.raw,grameo::Span {start:0,end:2});assert_eq!(metadata.normalization,NEEDS_GLOBAL_BINDING);
        assert!(matches!(builder.arena.get(original[1]),Value::Recipe {recipe:ContextRecipe::Atom {raw:grameo::Span {start:5,end:8},..},..}));
    }

    #[test]
    #[ignore="explicit contextual Pratt cache census against local WordPress sources"]
    fn profile_contextual_pratt_census() {
        let root=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../examples/webroot/wordpress/wp-includes");
        for name in ["post.php","functions.php","blocks/blocks-json.php","deprecated.php","general-template.php"] {
            CONTEXT_CENSUS.with(|cell|cell.set([0;7]));
            let source=std::fs::read_to_string(root.join(name)).unwrap();let accepted=super::super::parse_direct(&source).is_ok();
            let count=CONTEXT_CENSUS.with(|cell|cell.get());
            println!("file={name} bytes={} accepted={accepted} cache_roots={} contextual_binary_roots={} recipe_operators={} captured_operand_leaves={} owned_operand_leaves={} retained_atom_reuses={} atom_recaptures={}",source.len(),count[0],count[1],count[2],count[3],count[4],count[5],count[6]);
        }
    }

}

fn declaration_rule(rule:usize)->bool {matches!(Rule::from_capture(grameo::CaptureRule::User(rule)),Some(Rule::function_declaration|Rule::namespace_statement|Rule::class_declaration|Rule::interface_declaration|Rule::trait_declaration|Rule::enum_declaration))}

impl Ast<'_>{
    fn finish_ast_rule(&mut self,rule:usize,span:grameo::Span,source:&str){
        if self.mode_enabled() { self.finish_rule(rule, span, source); }
        if rule == Rule::comparison_expression as usize || rule == Rule::bit_or_expression as usize || rule == Rule::logic_or_expression as usize { self.pop_mode(); }
    }
}

impl Ast<'_> {
 pub(super) fn enable_source_intents(&mut self, bindings:Vec<grameo::source_units::Binding<'static>>) {
  assert!(self.source_hooks.is_none());
  // An unbound optimization annotation keeps the ordinary AST builder.
  if bindings.is_empty(){return;}
  let root=bindings[0].rule;let target=bindings[0].target;
  assert!(bindings.iter().all(|b|b.rule==root&&b.target==target));
  let deferred=[Rule::function_declaration,Rule::method_declaration,Rule::closure_expression,Rule::arrow_function].map(|r|r as usize);
  let mut observed=vec![root,target];observed.extend(bindings.iter().filter_map(|b|b.selector));observed.extend(deferred);observed.sort_unstable();observed.dedup();
  let state=self.arena.alloc(Value::SourceState(SourceState{head:None,deferred:0,last_intent:None}));
  self.source_hooks=Some(SourceHooks{state,bindings,root,target,observed,deferred});
 }
 fn rollback_source_pratt(&mut self,mark:usize,span:grameo::Span){
  let Some(hooks)=&self.source_hooks else{self.rollback_mark(mark);return};
  let state_id=hooks.state;
  let Value::SourceState(state)=*self.arena.get(state_id)else{unreachable!()};
  let mut head=state.last_intent;let mut retained=Vec::new();
  while let Some(id)=head{
   let Value::SourceIntent {request,previous,creation_mark}=self.arena.get(id)else{unreachable!()};
   if *creation_mark<mark{break;}
   assert!(request.span.start>=span.start&&request.span.end<=span.end,"source intent outside completed Pratt island");
   retained.push(request.clone());head=*previous;
  }
  self.rollback_mark(mark);
  if retained.is_empty(){return;}
  let Value::SourceState(mut state)=*self.arena.get(state_id)else{unreachable!()};
  for request in retained.into_iter().rev(){
   let creation_mark=self.arena.checkpoint()+self.mode_undo.len();
   state.last_intent=Some(self.arena.alloc(Value::SourceIntent {request,previous:state.last_intent,creation_mark}));
  }
  self.edit_source_state(state_id,state);
 }
 pub(super) fn take_source_intents(&mut self)->Vec<grameo::source_units::Request<String>>{std::mem::take(&mut self.source_intents)}
 fn edit_source_state(&mut self,id:Id,next:SourceState){self.arena.edit(id,|value|{let Value::SourceState(old)=value else{unreachable!()};MetadataUndo::SourceState(std::mem::replace(old,next))},|value,undo|{let(Value::SourceState(state),MetadataUndo::SourceState(old))=(value,undo)else{unreachable!()};*state=old;});}
 fn edit_source_frame(&mut self,id:Id,next:SourceFrame){self.arena.edit(id,|value|{let Value::SourceFrame(old)=value else{unreachable!()};MetadataUndo::SourceFrame(std::mem::replace(old,next))},|value,undo|{let(Value::SourceFrame(frame),MetadataUndo::SourceFrame(old))=(value,undo)else{unreachable!()};*frame=old;});}
 fn source_begin(&mut self,rule:usize,start:usize){
  let Some(h)=&self.source_hooks else{return};let id=h.state;let Value::SourceState(mut state)=*self.arena.get(id)else{unreachable!()};let root=h.root;let deferred=h.deferred.contains(&rule);
  if deferred{state.deferred+=1;}
  if rule==root{state.head=Some(self.arena.alloc(Value::SourceFrame(SourceFrame{parent:state.head,start,selector:None,target:None,targets:0})));}
  if deferred||rule==root{self.edit_source_state(id,state);}
 }
 fn source_finish(&mut self,rule:usize,span:grameo::Span,source:&str)->Result<(),grameo::builder::BuildError>{
  let Some(h)=&self.source_hooks else{return Ok(())};let id=h.state;let root=h.root;let target_rule=h.target;let deferred=h.deferred.contains(&rule);let selector=h.bindings.iter().any(|b|b.selector==Some(rule));
  let Value::SourceState(mut state)=*self.arena.get(id)else{unreachable!()};let fail=|message:&str|grameo::builder::BuildError{span,message:message.into()};
  if rule==root{
   let head=state.head.ok_or_else(||fail("missing source-unit frame"))?;let Value::SourceFrame(frame)=*self.arena.get(head)else{unreachable!()};
   if frame.start!=span.start{return Err(fail("mismatched source-unit frame"));}
   let binding=h.bindings.iter().find(|b|b.selector==frame.selector).copied().ok_or_else(||fail("missing source-unit selector binding"))?;
   if frame.targets!=1{return Err(fail("source-unit requires one semantic target"));}let target=frame.target.unwrap();
   let mut request=binding.request(source[target.start..target.end].to_owned(),span);
   request.preparation=if state.deferred==0{grameo::source_units::PreparationContext::Immediate}else{grameo::source_units::PreparationContext::Deferred};
   let creation_mark=self.arena.checkpoint()+self.mode_undo.len();
   state.last_intent=Some(self.arena.alloc(Value::SourceIntent {request,previous:state.last_intent,creation_mark}));state.head=frame.parent;self.edit_source_state(id,state);
  }
  if let Some(head)=state.head{let Value::SourceFrame(mut frame)=*self.arena.get(head)else{unreachable!()};let mut changed=false;
   if selector{if frame.selector.is_some(){return Err(fail("ambiguous source-unit selector"));}frame.selector=Some(rule);changed=true;}
   if rule==target_rule{frame.target=Some(span);frame.targets+=1;changed=true;}
   if changed{self.edit_source_frame(head,frame);}
  }
  if deferred{state.deferred=state.deferred.checked_sub(1).ok_or_else(||fail("unbalanced deferred source scope"))?;self.edit_source_state(id,state);}
  Ok(())
 }
}

#[cfg(test)]
mod eligibility_cache_tests {
    use super::*;

    #[test]
    fn repeated_admission_matches_pure_predicate_across_offsets_and_modes() {
        let source = "$a + 2 * 3; $b->field; 'text' . \"tail\"; $c ?? $d;";
        let mut ast = Ast::new(source);
        for mode in [false, true, false] {
            ast.pratt_only = mode;
            for offset in 0..=source.len() {
                let expected = ast.pratt_eligible(offset, source);
                assert_eq!(ast.cached_pratt_eligible(offset), expected);
                assert_eq!(ast.cached_pratt_eligible(offset), expected);
            }
        }
    }

    #[test]
    fn public_frontend_reuses_admission_and_preserves_complete_peg_ast() {
        let source = "<?php $a = 2; $b = 3; $x = $a + $b * 4; $y = $x << 2; return $y;";
        ELIGIBILITY_CENSUS.with(|c| c.set((0, 0)));
        let direct = crate::parse(source).unwrap();
        let (scans, hits) = ELIGIBILITY_CENSUS.with(|c| c.get());
        assert!(scans > 0 && hits > 0, "active public path: scans={scans} hits={hits}");
        ELIGIBILITY_CENSUS.with(|c| c.set((0, 0)));
        let peg = crate::parse_with_peg(source).unwrap();
        assert_eq!(ELIGIBILITY_CENSUS.with(|c| c.get()), (0, 0), "PEG reference does not select optimized admission");
        assert_eq!(format!("{direct:?}"), format!("{peg:?}"));
    }
}

/// Accepted cache entries actually consumed by the contextual statement walker.
pub(super) fn consumed_expression_recipes(cache: &ExpressionCache) -> usize {
    cache.values().filter(|entry| matches!(entry,
        CachedExpression::String { first: None, .. }
        | CachedExpression::Ident { first: None, .. }
        | CachedExpression::Number { first: None, .. }
        | CachedExpression::Pending(None)
        | CachedExpression::Deferred { first: None, .. }
        | CachedExpression::Contextual(None)
        | CachedExpression::Owned(_))).count()
}
