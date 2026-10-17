//! Grammar-finished C syntax ASTs and ordered body plans; typed lowering stays
//! at the source-order consumer barrier and contextual syntax keeps captures.
use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use grameo::{Span, arena::EditArena, builder::Builder, compat::Pair, tree::CaptureHandle};
use vybe_ast::{Argument, ArrayElement, BinOp, BreakTarget, ContinueTarget, ExprKind, Expression, Param, Statement, StmtKind, SwitchCase};

use crate::Rule;
use super::owned::{PlainDeclaratorSyntax, PlainScalarDeclarator, construct_plain_declarator};

#[derive(Clone, Copy)]
struct Origin {
    anchor: Option<Span>,
}

enum Recipe {
    Literal { rule: Rule, span: Span },
    Ident(Span),
    Prefix { op: &'static str, operand: Box<Recipe> },
    Array(Vec<Recipe>),
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(crate) struct ReplayWork {
    pub(crate) literals: usize,
    pub(crate) prefixes: usize,
    pub(crate) arrays: usize,
}

thread_local! {
    static REPLAY_OBSERVER: std::cell::Cell<Option<fn(ReplayWork)>> = const { std::cell::Cell::new(None) };
}

pub(crate) fn with_replay_observer<T>(observer: fn(ReplayWork), action: impl FnOnce() -> T) -> T {
    struct Restore(Option<fn(ReplayWork)>);
    impl Drop for Restore {
        fn drop(&mut self) { REPLAY_OBSERVER.with(|slot| slot.set(self.0)); }
    }
    let _restore = Restore(REPLAY_OBSERVER.with(|slot| slot.replace(Some(observer))));
    action()
}

impl Recipe {
    fn replay(&self, source: &str) -> Expression {
        if let Some(observer) = REPLAY_OBSERVER.with(|slot| slot.get()) {
            let work = self.replay_work();
            let value = self.replay_unprofiled(source);
            observer(work);
            value
        } else { self.replay_unprofiled(source) }
    }

    fn replay_work(&self) -> ReplayWork {
        match self {
            Self::Ident(_) => ReplayWork::default(),
            Self::Literal { .. } => ReplayWork { literals: 1, ..Default::default() },
            Self::Prefix { operand, .. } => {
                let mut work = operand.replay_work();
                work.prefixes += 1;
                work
            }
            Self::Array(values) => values.iter().fold(ReplayWork { arrays: 1, ..Default::default() }, |mut total, value| {
                let work = value.replay_work();
                total.literals += work.literals;
                total.prefixes += work.prefixes;
                total.arrays += work.arrays;
                total
            }),
        }
    }

    fn replay_unprofiled(&self, source: &str) -> Expression {
        match self {
            Self::Ident(span) => Expression::new(ExprKind::Ident(source[span.start..span.end].to_string())),
            Self::Literal { rule, span } => {
                crate::walker::construct_literal(*rule, &source[span.start..span.end])
            }
            Self::Array(values) => array(values.iter().map(|value| value.replay_unprofiled(source))),
            Self::Prefix { op, operand } =>
                crate::walker::construct_numeric_prefix(op, operand.replay_unprofiled(source)),
        }
    }

    fn numeric(&self) -> bool {
        match self {
            Self::Literal { rule, .. } => !matches!(rule, Rule::string_literal | Rule::wide_string_literal),
            Self::Prefix { operand, .. } => operand.numeric(),
            Self::Array(_) | Self::Ident(_) => false,
        }
    }
}

fn array(values: impl IntoIterator<Item = Expression>) -> Expression {
    Expression::new(ExprKind::Array(values.into_iter().map(|value| ArrayElement {
        key: None, value, spread: false, by_ref: false,
    }).collect()))
}

/// A complete scope-independent statement, or the C empty statement (no node).
pub(crate) enum OwnedStatement {
    Emit(Statement),
    Empty,
}

/// Ordered function/block children. Contextual children retain their original
/// capture anchors; entering/leaving C scopes remains the caller's responsibility.
pub(crate) enum BodyItem<'a> {
    Owned { pair: Pair<'a, Rule>, value: OwnedStatement },
    Contextual(Pair<'a, Rule>),
}

#[derive(Clone, Copy)]
enum ControlRecipe {
    Break,
    Continue(Option<Span>),
    Goto(Span),
    ReturnVoid,
    Empty,
}

impl ControlRecipe {
    fn replay(self, source: &str) -> OwnedStatement {
        let kind = match self {
            Self::Break => StmtKind::Break(BreakTarget::Implicit),
            Self::Continue(label) => StmtKind::Continue(match label {
                Some(span) => ContinueTarget::Label(source[span.start..span.end].to_string()),
                None => ContinueTarget::Implicit,
            }),
            Self::Goto(span) => StmtKind::GoTo(source[span.start..span.end].to_string()),
            Self::ReturnVoid => StmtKind::Return(None),
            Self::Empty => return OwnedStatement::Empty,
        };
        OwnedStatement::Emit(vybe_platform_libc::emitter::build::stmt(kind))
    }
}

struct Control {
    value: OwnedStatement,
    recipe: ControlRecipe,
}

struct BodyReceipt {
    capture: CaptureHandle,
    span: Span,
    control: Option<Control>,
}

struct CachedStatement {
    first: Option<OwnedStatement>,
    recipe: ControlRecipe,
}

enum SyntaxRecipe {
    Primary { span: Span, recipe: Recipe },
    Unary { span: Span, op: Option<&'static str>, operand: Rc<SyntaxRecipe>, capture: Option<CaptureHandle> },
    Sequence { span: Span, operands: Vec<Rc<SyntaxRecipe>> },
    Initializer { span: Span, operand: Rc<SyntaxRecipe> },
    Grouped { span: Span, operand: Rc<SyntaxRecipe> },
}

impl SyntaxRecipe {
    fn rule_span(&self) -> (Rule, Span) {
        match self {
            Self::Primary { span, .. } => (Rule::primary_expression, *span),
            Self::Unary { span, .. } => (Rule::unary_expression, *span),
            Self::Sequence { span, .. } => (Rule::expression, *span),
            Self::Initializer { span, .. } => (Rule::initializer, *span),
            Self::Grouped { span, .. } => (Rule::primary_expression, *span),
        }
    }

    fn replay(&self, source: &str) -> Expression {
        match self {
            Self::Primary { recipe, .. } => recipe.replay(source),
            Self::Initializer { operand, .. } | Self::Grouped { operand, .. } => operand.replay(source),
            Self::Unary { op, operand, .. } => {
                let value = operand.replay(source);
                if let Some(op) = op { crate::walker::construct_numeric_prefix(op, value) } else { value }
            }
            Self::Sequence { operands, .. } => {
                let mut values: Vec<_> = operands.iter().map(|operand| operand.replay(source)).collect();
                if values.len() == 1 { values.pop().unwrap() }
                else { Expression::new(ExprKind::Sequence(values)) }
            }
        }
    }
}

struct CachedSyntax {
    first: Option<Expression>,
    recipe: Rc<SyntaxRecipe>,
}

enum BinaryRecipe {
    Atom { capture: Option<CaptureHandle>, syntax: Rc<SyntaxRecipe> },
    Infix { op: BinOp, left: Rc<BinaryRecipe>, right: Rc<BinaryRecipe> },
}

impl BinaryRecipe {
    fn anchors<'a>(&self, pair: &Pair<'a, Rule>) -> Option<Vec<Pair<'a, Rule>>> {
        let mut pending = vec![self];
        let mut anchors = Vec::new();
        while let Some(recipe) = pending.pop() {
            match recipe {
                Self::Atom { capture, .. } => {
                    let anchor = pair.resolve_capture((*capture)?)?;
                    if anchor.as_rule() != Rule::unary_expression { return None; }
                    anchors.push(anchor);
                }
                Self::Infix { left, right, .. } => { pending.push(right); pending.push(left); }
            }
        }
        Some(anchors)
    }

    fn replay(&self, source: &str) -> Expression {
        let mut pending = vec![(self, false)];
        let mut values = Vec::new();
        while let Some((recipe, reduced)) = pending.pop() {
            match recipe {
                Self::Atom { syntax, .. } => values.push(syntax.replay(source)),
                Self::Infix { op, left, right } => {
                    if !reduced {
                        pending.push((recipe, true)); pending.push((right, false)); pending.push((left, false));
                    } else {
                        let right = values.pop().unwrap();
                        let left = values.pop().unwrap();
                        values.push(Expression::new(ExprKind::Binary { op: *op, left: Box::new(left), right: Box::new(right) }));
                    }
                }
            }
        }
        values.pop().unwrap()
    }
}

struct CachedBinary { first: Option<Expression>, recipe: Rc<BinaryRecipe> }

/// One original binary grammar region, never a generic recursive C rewrite.
/// All capture anchors have been validated before this value reaches a caller.
pub(crate) struct BinaryRegion<'a> {
    value: Expression,
    recipe: Rc<BinaryRecipe>,
    operands: Vec<Pair<'a, Rule>>,
}

impl<'a> BinaryRegion<'a> {
    pub(crate) fn operand_count(&self) -> usize { self.operands.len() }

    /// Normalize each original unary operand in source order, including the
    /// caller's integral-value operand policy, then return the completed region.
    /// The caller must run its binary-region rewrite exactly once afterward.
    pub(crate) fn lower(self, mut operand: impl FnMut(Pair<'a, Rule>, Expression) -> Expression) -> Expression {
        enum Task { Visit(Rc<BinaryRecipe>, Expression), Finish(Expression) }
        let mut anchors = self.operands.into_iter();
        let mut pending = vec![Task::Visit(self.recipe, self.value)];
        let mut values = Vec::new();
        while let Some(task) = pending.pop() {
            match task {
                Task::Visit(recipe, mut value) => match &*recipe {
                    BinaryRecipe::Atom { .. } => values.push(operand(anchors.next().unwrap(), value)),
                    BinaryRecipe::Infix { op, left: lrecipe, right: rrecipe } => {
                        let ExprKind::Binary { op: actual, left, right } = &mut value.kind else { unreachable!("owned binary shape"); };
                        debug_assert_eq!(actual, op);
                        // Retain the grammar-built boxes while the owned children
                        // pass through contextual lowering. No clone or TLS lease.
                        let lvalue = std::mem::replace(left.as_mut(), Expression::new(ExprKind::Lit(vybe_ast::Literal::Null)));
                        let rvalue = std::mem::replace(right.as_mut(), Expression::new(ExprKind::Lit(vybe_ast::Literal::Null)));
                        pending.push(Task::Finish(value));
                        pending.push(Task::Visit(rrecipe.clone(), rvalue));
                        pending.push(Task::Visit(lrecipe.clone(), lvalue));
                    }
                },
                Task::Finish(mut value) => {
                    let ExprKind::Binary { left, right, .. } = &mut value.kind else { unreachable!(); };
                    *right.as_mut() = values.pop().unwrap();
                    *left.as_mut() = values.pop().unwrap();
                    values.push(super::owned::finish_binary_syntax(value));
                }
            }
        }
        debug_assert!(anchors.next().is_none());
        values.pop().unwrap()
    }
}

#[derive(Clone, Copy)]
enum BinaryOrigin { Syntax, Binary(Option<Span>) }

struct CachedDeclarator {
    first: Option<PlainScalarDeclarator>,
    syntax: PlainDeclaratorSyntax,
    initializer: Option<Rc<SyntaxRecipe>>,
}

impl CachedDeclarator {
    fn take(&mut self, source: &str) -> PlainScalarDeclarator {
        self.first.take().unwrap_or_else(|| construct_plain_declarator(
            source[self.syntax.declarator.start..self.syntax.declarator.end].to_string(),
            self.initializer.as_ref().map(|recipe| recipe.replay(source)), self.syntax,
        ))
    }
}

#[derive(Clone, Copy)]
enum SyntaxOrigin { Primary, Syntax }

/// An ordinary named parameter's syntax, not its resolved C type or ABI.
/// The ordered consumer supplies the final hint/access or retains fallback.
pub(crate) struct PlainParameter<'a> {
    pub(crate) parameter: Param,
    pub(crate) specifiers: Pair<'a, Rule>,
    pub(crate) declaration_span: Span,
    pub(crate) declarator_span: Span,
}

#[derive(Clone, Copy)]
struct ParameterRecipe {
    name: Span,
    specifiers: CaptureHandle,
    specifier_span: Span,
    declarator_span: Span,
}

struct CachedParameter { first: Option<Param>, recipe: ParameterRecipe }

/// Completed grammar structure, with contextual work deliberately unperformed.
/// `finish` accepts only the ordered consumer's already-normalized values.
pub(crate) struct IfRegion<'a> {
    pub(crate) condition: Pair<'a, Rule>,
    pub(crate) then_branch: Pair<'a, Rule>,
    pub(crate) else_branch: Option<Pair<'a, Rule>>,
    value: Statement,
}

impl IfRegion<'_> {
    pub(crate) fn finish(mut self, cond: Expression, then_body: Vec<Statement>,
        else_body: Option<Vec<Statement>>) -> Statement
    {
        let StmtKind::If { cond: target, then_body: then_target, else_body: else_target, .. }
            = &mut self.value.kind else { unreachable!("direct if constructor"); };
        *target = cond;
        *then_target = then_body;
        *else_target = else_body;
        self.value
    }
}

#[derive(Clone, Copy)]
struct ContextAnchor { rule: Rule, span: Span, capture: CaptureHandle }

impl ContextAnchor {
    fn resolve<'a>(self, pair: &Pair<'a, Rule>) -> Option<Pair<'a, Rule>> {
        let anchor = pair.resolve_capture(self.capture)?;
        let span = anchor.as_span();
        (anchor.as_rule() == self.rule && span.start() == self.span.start
            && span.end() == self.span.end).then_some(anchor)
    }
}

#[derive(Clone, Copy)]
struct IfRecipe { condition: ContextAnchor, then_branch: ContextAnchor, else_branch: Option<ContextAnchor> }

fn if_shell() -> Statement {
    // This private shell cannot escape before every contextual field is filled.
    vybe_platform_libc::emitter::build::stmt(StmtKind::If {
        cond: Expression::new(ExprKind::Lit(vybe_ast::Literal::Null)), then_body: Vec::new(),
        elifs: Vec::new(), else_body: None,
    })
}

struct CachedIf { first: Option<Statement>, recipe: IfRecipe }

pub(crate) struct WhileRegion<'a> {
    pub(crate) condition: Pair<'a, Rule>,
    pub(crate) body: Pair<'a, Rule>,
    value: Statement,
}

impl WhileRegion<'_> {
    pub(crate) fn finish(mut self, cond: Expression, body: Vec<Statement>) -> Statement {
        let StmtKind::While { cond: target, body: body_target, .. } = &mut self.value.kind
            else { unreachable!("direct while constructor"); };
        *target = cond;
        *body_target = body;
        self.value
    }
}

pub(crate) struct DoWhileRegion<'a> {
    pub(crate) condition: Pair<'a, Rule>,
    pub(crate) body: Pair<'a, Rule>,
    value: Statement,
}

impl DoWhileRegion<'_> {
    pub(crate) fn finish(mut self, cond: Expression, body: Vec<Statement>) -> Statement {
        let StmtKind::DoWhile { cond: target, body: body_target, .. } = &mut self.value.kind
            else { unreachable!("direct do-while constructor"); };
        *target = cond;
        *body_target = body;
        self.value
    }
}

#[derive(Clone, Copy)]
struct LoopRecipe { condition: ContextAnchor, body: ContextAnchor, do_while: bool }

impl LoopRecipe {
    fn shell(self) -> Statement {
        let cond = Expression::new(ExprKind::Lit(vybe_ast::Literal::Null));
        let kind = if self.do_while {
            StmtKind::DoWhile { body: Vec::new(), cond, until: false }
        } else { StmtKind::While { cond, body: Vec::new(), else_body: None } };
        vybe_platform_libc::emitter::build::stmt(kind)
    }
}

struct CachedLoop { first: Option<Statement>, recipe: LoopRecipe }

/// Original clause boundaries, including empty init/condition wrappers. The
/// ordered consumer owns declaration scope and every clause-specific rewrite.
pub(crate) struct ForRegion<'a> {
    pub(crate) init: Pair<'a, Rule>,
    pub(crate) condition: Pair<'a, Rule>,
    pub(crate) update: Option<Pair<'a, Rule>>,
    pub(crate) body: Pair<'a, Rule>,
    value: Statement,
}

impl ForRegion<'_> {
    pub(crate) fn finish(mut self, init: Option<Box<Statement>>, cond: Option<Expression>,
        update: Option<Expression>, body: Vec<Statement>) -> Statement
    {
        let StmtKind::For { init: init_target, cond: cond_target,
            update: update_target, body: body_target } = &mut self.value.kind
            else { unreachable!("direct for constructor"); };
        *init_target = init;
        *cond_target = cond;
        *update_target = update;
        *body_target = body;
        self.value
    }
}

#[derive(Clone, Copy)]
struct ForRecipe {
    init: ContextAnchor,
    condition: ContextAnchor,
    update: Option<ContextAnchor>,
    body: ContextAnchor,
}

fn for_shell() -> Statement {
    vybe_platform_libc::emitter::build::stmt(StmtKind::For {
        init: None, cond: None, update: None, body: Vec::new(),
    })
}

struct CachedFor { first: Option<Statement>, recipe: ForRecipe }

/// Case collection, selector promotion and fallthrough stay consumer-owned.
pub(crate) struct SwitchRegion<'a> {
    pub(crate) subject: Pair<'a, Rule>,
    pub(crate) body: Pair<'a, Rule>,
    value: Statement,
}

impl SwitchRegion<'_> {
    pub(crate) fn finish(mut self, expr: Expression, cases: Vec<SwitchCase>,
        default: Option<Vec<Statement>>) -> Statement
    {
        let StmtKind::Switch { expr: target, cases: cases_target, default: default_target }
            = &mut self.value.kind else { unreachable!("direct switch constructor"); };
        *target = expr;
        *cases_target = cases;
        *default_target = default;
        self.value
    }
}

#[derive(Clone, Copy)]
struct SwitchRecipe { subject: ContextAnchor, body: ContextAnchor }

fn switch_shell() -> Statement {
    vybe_platform_libc::emitter::build::stmt(StmtKind::Switch {
        expr: Expression::new(ExprKind::Lit(vybe_ast::Literal::Null)),
        cases: Vec::new(), default: None,
    })
}

struct CachedSwitch { first: Option<Statement>, recipe: SwitchRecipe }

/// A value-bearing return. Expression typing, character promotion and function
/// return normalization belong to the ordered consumer, not the matcher.
pub(crate) struct ReturnRegion<'a> {
    pub(crate) expression: Pair<'a, Rule>,
    value: Statement,
}

impl ReturnRegion<'_> {
    pub(crate) fn finish(mut self, value: Expression) -> Statement {
        let StmtKind::Return(target) = &mut self.value.kind
            else { unreachable!("direct return constructor"); };
        *target = Some(value);
        self.value
    }
}

fn return_shell() -> Statement {
    vybe_platform_libc::emitter::build::stmt(StmtKind::Return(None))
}

struct CachedReturn { first: Option<Statement>, expression: ContextAnchor }

/// The ordinary call constructor only. C argument modes, ABI and adapter
/// selection must finish in the ordered consumer before this node is filled.
pub(crate) struct CallRegion<'a> {
    pub(crate) arguments: Vec<Pair<'a, Rule>>,
    value: Expression,
}

impl CallRegion<'_> {
    pub(crate) fn finish(mut self, callee: Expression, args: Vec<Argument>) -> Expression {
        let ExprKind::Call { callee: target, args: args_target, .. } = &mut self.value.kind
            else { unreachable!("direct call constructor"); };
        **target = callee;
        *args_target = args;
        self.value
    }
}

fn call_shell() -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lit(vybe_ast::Literal::Null))),
        args: Vec::new(), optional: false,
    })
}

struct CachedCall { first: Option<Expression>, arguments: Vec<ContextAnchor> }

pub(crate) struct IndexRegion<'a> {
    pub(crate) index: Pair<'a, Rule>,
    value: Expression,
}

impl IndexRegion<'_> {
    pub(crate) fn finish(mut self, object: Expression, index: Expression) -> Expression {
        let ExprKind::Index { object: target, index: index_target, .. } = &mut self.value.kind
            else { unreachable!("direct index constructor"); };
        **target = object;
        **index_target = index;
        self.value
    }
}

pub(crate) struct MemberRegion<'a> {
    pub(crate) field: Pair<'a, Rule>,
    value: Expression,
}

impl MemberRegion<'_> {
    pub(crate) fn finish(mut self, object: Expression) -> Expression {
        let ExprKind::Member { object: target, .. } = &mut self.value.kind
            else { unreachable!("direct member constructor"); };
        **target = object;
        self.value
    }
}

fn index_shell() -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(Expression::new(ExprKind::Lit(vybe_ast::Literal::Null))),
        index: Box::new(Expression::new(ExprKind::Lit(vybe_ast::Literal::Null))), null_safe: false,
    })
}

fn member_shell(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::Lit(vybe_ast::Literal::Null))),
        field: field.into(), null_safe: false,
    })
}

struct CachedPlace { first: Option<Expression>, anchor: ContextAnchor }

pub(crate) struct FunctionRegion<'a> {
    pub(crate) specifiers: Pair<'a, Rule>,
    pub(crate) declarator: Pair<'a, Rule>,
    pub(crate) declarations: Vec<Pair<'a, Rule>>,
    pub(crate) body: Pair<'a, Rule>,
    value: Statement,
}

impl<'a> FunctionRegion<'a> {
    pub(crate) fn parts(&self) -> impl Iterator<Item = Pair<'a, Rule>> + '_ {
        std::iter::once(self.specifiers.clone()).chain(std::iter::once(self.declarator.clone()))
            .chain(self.declarations.iter().cloned()).chain(std::iter::once(self.body.clone()))
    }

    pub(crate) fn finish(mut self, name: String, params: Vec<Param>,
        return_type: Option<String>, body: Vec<Statement>) -> Statement
    {
        let StmtKind::FunctionDecl { name: name_target, params: params_target,
            return_type: return_target, body: body_target, .. } = &mut self.value.kind
            else { unreachable!("direct function constructor"); };
        *name_target = name;
        *params_target = params;
        *return_target = return_type;
        *body_target = body;
        self.value
    }
}

fn function_shell() -> Statement {
    crate::walker::construct_function_declaration(String::new(), Vec::new(), None, Vec::new())
}

struct CachedFunction { first: Option<Statement>, parts: Vec<ContextAnchor> }

enum Node {
    #[cfg(test)]
    Frame { rule: usize, start: usize },
    Ready { origin: Origin, value: Expression, recipe: Recipe },
    Primary { span: Span, value: Expression, recipe: Recipe },
    Syntax { value: Expression, recipe: Rc<SyntaxRecipe> },
    PlainName { span: Span, name: Span },
    Declarator { span: Span, value: PlainScalarDeclarator, initializer: Option<Rc<SyntaxRecipe>> },
    Binary { span: Option<Span>, value: Expression, recipe: Rc<BinaryRecipe> },
    ParameterSpecifiers { span: Span, capture: Option<CaptureHandle> },
    Parameter { span: Span, value: Param, recipe: ParameterRecipe },
    ContextAnchor(ContextAnchor),
    IfRegion { span: Span, value: Statement, recipe: IfRecipe },
    LoopRegion { span: Span, value: Statement, recipe: LoopRecipe },
    ForRegion { span: Span, value: Statement, recipe: ForRecipe },
    SwitchRegion { span: Span, value: Statement, recipe: SwitchRecipe },
    ReturnRegion { span: Span, value: Statement, expression: ContextAnchor },
    CallArguments(Vec<ContextAnchor>),
    CallRegion { span: Span, value: Expression, arguments: Vec<ContextAnchor> },
    IndexRegion { span: Span, value: Expression, anchor: ContextAnchor },
    MemberRegion { rule: Rule, span: Span, value: Expression, anchor: ContextAnchor },
    FunctionRegion { span: Span, value: Statement, parts: Vec<ContextAnchor> },
    Prefix(&'static str),
    Label(Span),
    Control(Control),
    BodyItem(BodyReceipt),
    Body { span: Span, items: Vec<BodyReceipt> },
    Unsupported,
}

enum Undo {
    Array(Vec<(usize, Span)>),
    Prefix(Origin),
    BodyItem,
    Label,
    Primary(Origin),
    Syntax(Vec<SyntaxOrigin>),
    PlainName,
    Declarator(Span),
    Binary([BinaryOrigin; 2]),
    BinarySpan(Option<Span>),
    Parameter(Span),
    #[cfg(test)]
    ReferenceArray(Vec<Origin>),
}

#[derive(Clone, Copy)]
struct Frame {
    rule: usize,
    start: usize,
    base: usize,
}

enum StateUndo {
    State { depth: usize, literal_active: bool },
    PopFrame,
    PushFrame(Frame),
}

type Input = (usize, usize);
thread_local! {
    static CACHE: RefCell<Option<HashMap<Input, Initializers>>> = const { RefCell::new(None) };
    #[cfg(test)]
    static REFERENCE_CONSTRUCTION: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

#[cfg(test)]
fn with_reference_construction<T>(reference: bool, action: impl FnOnce() -> T) -> T {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { REFERENCE_CONSTRUCTION.with(|slot| slot.set(self.0)); }
    }
    let _restore = Restore(REFERENCE_CONSTRUCTION.with(|slot| slot.replace(reference)));
    action()
}

/// Enclose matching AND ordered lowering, restoring nested scopes on unwind.
pub(crate) fn with_cache_scope<T>(action: impl FnOnce() -> T) -> T {
    struct Restore(Option<HashMap<Input, Initializers>>);
    impl Drop for Restore {
        fn drop(&mut self) { CACHE.with(|cache| *cache.borrow_mut() = self.0.take()); }
    }
    let _restore = Restore(CACHE.with(|cache| cache.replace(Some(HashMap::new()))));
    action()
}

/// Reparse invalidation precedes matching, including failed matches. No scope
/// means no publication or elision; callers without the handoff keep captures.
pub(crate) fn build_captures<T, E>(
    source: &str, elide: bool,
    build: impl FnOnce(&mut Direct<'_>) -> Result<T, E>,
) -> Result<T, E> {
    let input = (source.as_ptr() as usize, source.len());
    let enabled = CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        if let Some(cache) = cache.as_mut() { cache.remove(&input); true } else { false }
    });
    let mut builder = if elide && enabled {
        Direct::with_capture_elision(source)
    } else { Direct::new(source) };
    let value = build(&mut builder)?;
    let initializers = builder.into_initializers();
    if enabled {
        CACHE.with(|cache| {
            if let Some(cache) = cache.borrow_mut().as_mut() { cache.insert(input, initializers); }
        });
    }
    Ok(value)
}

/// Start a new full-source unit sequence, invalidating older parse generations.
pub(crate) fn reset_unit_captures(source: &str) {
    CACHE.with(|cache| {
        if let Some(cache) = cache.borrow_mut().as_mut() {
            cache.remove(&(source.as_ptr() as usize, source.len()));
        }
    });
}

/// One original-grammar unit match over the full source at an absolute offset.
/// Failed/uncommitted units publish nothing and preserve earlier committed ASTs.
pub(crate) fn build_unit_captures<T, E>(
    source: &str, elide: bool,
    build: impl FnOnce(&mut Direct<'_>) -> Result<T, E>,
) -> Result<T, E> {
    let enabled = CACHE.with(|cache| cache.borrow().is_some());
    let mut builder = if enabled && elide { Direct::with_capture_elision(source) }
        else { Direct::new(source) };
    let value = build(&mut builder)?;
    if enabled && builder.committed {
        let input = (source.as_ptr() as usize, source.len());
        let additions = builder.into_initializers();
        CACHE.with(|cache| {
            if let Some(cache) = cache.borrow_mut().as_mut() {
                if let Some(existing) = cache.get_mut(&input) { existing.append(additions); }
                else { cache.insert(input, additions); }
            }
        });
    }
    Ok(value)
}

pub(crate) fn take_initializer_list(pair: &Pair<'_, Rule>) -> Option<Expression> {
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| cache.borrow_mut().as_mut()?
        .get_mut(&(source.as_ptr() as usize, source.len()))?.take(pair))
}

/// Syntax only: identifiers are unresolved. The ordered consumer must apply
/// its ordinary primary-name normalization before using this expression.
pub(crate) fn take_primary_expression(pair: &Pair<'_, Rule>) -> Option<Expression> {
    if pair.as_rule() != Rule::primary_expression { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        if let Some(entry) = values.primaries.get_mut(&(span.start(), span.end())) {
            Some(entry.first.take().unwrap_or_else(|| entry.recipe.replay(source)))
        } else {
            // This API's consumer normalizes a single name, not a compound
            // expression. Grouping is consumed only by the syntax-tree API.
            let entry = values.syntax.get(&(Rule::primary_expression, span.start(), span.end()))?;
            if !matches!(&*entry.recipe, SyntaxRecipe::Primary { .. }) { return None; }
            values.take_syntax(pair)
        }
    })
}

/// Grammar-built syntax, not typed C lowering. Supported trees contain only
/// primary values, numeric +/-/~ prefixes, and comma sequences. Resolve names
/// in source order at the consumer barrier; do not treat Ident as lowered C.
pub(crate) fn take_syntax_expression(pair: &Pair<'_, Rule>) -> Option<Expression> {
    if !matches!(pair.as_rule(), Rule::expression | Rule::unary_expression) { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| cache.borrow_mut().as_mut()?
        .get_mut(&(source.as_ptr() as usize, source.len()))?.take_syntax(pair))
}

/// Raw syntax, unlike policy-applied Owned/Mixed receipts. The caller must
/// retain its scalar-admission guard and register the binding before normalizing
/// this initializer. No storage/type/default-value decision is made here.
pub(crate) fn take_plain_declarator(pair: &Pair<'_, Rule>) -> Option<PlainScalarDeclarator> {
    if pair.as_rule() != Rule::init_declarator { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        Some(values.declarators.get_mut(&(span.start(), span.end()))?.take(source))
    })
}

pub(crate) fn take_binary_region<'a>(pair: &Pair<'a, Rule>) -> Option<BinaryRegion<'a>> {
    if pair.as_rule() != Rule::binary_expression { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.binaries.get_mut(&(span.start(), span.end()))?;
        let operands = entry.recipe.anchors(pair)?;
        Some(BinaryRegion { value: entry.first.take().unwrap_or_else(|| entry.recipe.replay(source)),
            recipe: entry.recipe.clone(), operands })
    })
}

pub(crate) fn take_plain_parameter<'a>(pair: &Pair<'a, Rule>) -> Option<PlainParameter<'a>> {
    if pair.as_rule() != Rule::parameter_decl { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.parameters.get_mut(&(span.start(), span.end()))?;
        let recipe = entry.recipe;
        // A stale/foreign capture must not consume any AST or run C policy.
        let specifiers = pair.resolve_capture(recipe.specifiers)?;
        let actual = specifiers.as_span();
        if specifiers.as_rule() != Rule::declaration_specifiers
            || actual.start() != recipe.specifier_span.start || actual.end() != recipe.specifier_span.end
        { return None; }
        let parameter = entry.first.take().unwrap_or_else(|| crate::walker::construct_plain_parameter(
            source[recipe.name.start..recipe.name.end].to_string()));
        Some(PlainParameter { parameter, specifiers,
            declaration_span: Span { start: span.start(), end: span.end() },
            declarator_span: recipe.declarator_span })
    })
}

pub(crate) fn take_if_region<'a>(pair: &Pair<'a, Rule>) -> Option<IfRegion<'a>> {
    if pair.as_rule() != Rule::if_statement { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.ifs.get_mut(&(span.start(), span.end()))?;
        // Validate all anchors before spending the shell. The caller may then
        // leave this TLS borrow and run arbitrary scope-sensitive C lowering.
        let condition = entry.recipe.condition.resolve(pair)?;
        let then_branch = entry.recipe.then_branch.resolve(pair)?;
        let else_branch = match entry.recipe.else_branch {
            Some(anchor) => Some(anchor.resolve(pair)?), None => None,
        };
        Some(IfRegion { condition, then_branch, else_branch,
            value: entry.first.take().unwrap_or_else(if_shell) })
    })
}

fn take_loop_region<'a>(pair: &Pair<'a, Rule>, do_while: bool)
    -> Option<(Pair<'a, Rule>, Pair<'a, Rule>, Statement)>
{
    let rule = if do_while { Rule::do_while_statement } else { Rule::while_statement };
    if pair.as_rule() != rule { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.loops.get_mut(&(rule, span.start(), span.end()))?;
        if entry.recipe.do_while != do_while { return None; }
        let condition = entry.recipe.condition.resolve(pair)?;
        let body = entry.recipe.body.resolve(pair)?;
        Some((condition, body, entry.first.take().unwrap_or_else(|| entry.recipe.shell())))
    })
}

pub(crate) fn take_while_region<'a>(pair: &Pair<'a, Rule>) -> Option<WhileRegion<'a>> {
    let (condition, body, value) = take_loop_region(pair, false)?;
    Some(WhileRegion { condition, body, value })
}

pub(crate) fn take_do_while_region<'a>(pair: &Pair<'a, Rule>) -> Option<DoWhileRegion<'a>> {
    let (condition, body, value) = take_loop_region(pair, true)?;
    Some(DoWhileRegion { condition, body, value })
}

pub(crate) fn take_for_region<'a>(pair: &Pair<'a, Rule>) -> Option<ForRegion<'a>> {
    if pair.as_rule() != Rule::for_statement { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.fors.get_mut(&(span.start(), span.end()))?;
        let init = entry.recipe.init.resolve(pair)?;
        let condition = entry.recipe.condition.resolve(pair)?;
        let update = match entry.recipe.update {
            Some(anchor) => Some(anchor.resolve(pair)?), None => None,
        };
        let body = entry.recipe.body.resolve(pair)?;
        Some(ForRegion { init, condition, update, body,
            value: entry.first.take().unwrap_or_else(for_shell) })
    })
}

pub(crate) fn take_switch_region<'a>(pair: &Pair<'a, Rule>) -> Option<SwitchRegion<'a>> {
    if pair.as_rule() != Rule::switch_statement { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.switches.get_mut(&(span.start(), span.end()))?;
        let subject = entry.recipe.subject.resolve(pair)?;
        let body = entry.recipe.body.resolve(pair)?;
        Some(SwitchRegion { subject, body,
            value: entry.first.take().unwrap_or_else(switch_shell) })
    })
}

/// Void returns retain the existing complete control-statement handoff. Every
/// contextual anchor is validated before spending the first owned constructor.
pub(crate) fn take_return_region<'a>(pair: &Pair<'a, Rule>) -> Option<ReturnRegion<'a>> {
    if pair.as_rule() != Rule::return_statement { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.returns.get_mut(&(span.start(), span.end()))?;
        let expression = entry.expression.resolve(pair)?;
        Some(ReturnRegion { expression, value: entry.first.take().unwrap_or_else(return_shell) })
    })
}

pub(crate) fn take_call_region<'a>(pair: &Pair<'a, Rule>) -> Option<CallRegion<'a>> {
    if pair.as_rule() != Rule::call_suffix { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.calls.get_mut(&(span.start(), span.end()))?;
        let arguments: Option<Vec<_>> = entry.arguments.iter().map(|anchor| anchor.resolve(pair)).collect();
        Some(CallRegion { arguments: arguments?, value: entry.first.take().unwrap_or_else(call_shell) })
    })
}

pub(crate) fn take_index_region<'a>(pair: &Pair<'a, Rule>) -> Option<IndexRegion<'a>> {
    if pair.as_rule() != Rule::index_suffix { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.indices.get_mut(&(span.start(), span.end()))?;
        let index = entry.anchor.resolve(pair)?;
        Some(IndexRegion { index, value: entry.first.take().unwrap_or_else(index_shell) })
    })
}

pub(crate) fn take_member_region<'a>(pair: &Pair<'a, Rule>) -> Option<MemberRegion<'a>> {
    if !matches!(pair.as_rule(), Rule::member_suffix | Rule::arrow_suffix) { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.members.get_mut(&(pair.as_rule(), span.start(), span.end()))?;
        let field = entry.anchor.resolve(pair)?;
        let value = entry.first.take().unwrap_or_else(|| member_shell(field.as_str()));
        Some(MemberRegion { field, value })
    })
}

pub(crate) fn take_function_region<'a>(pair: &Pair<'a, Rule>) -> Option<FunctionRegion<'a>> {
    if pair.as_rule() != Rule::function_definition { return None; }
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&(source.as_ptr() as usize, source.len()))?;
        let entry = values.functions.get_mut(&(span.start(), span.end()))?;
        let parts: Option<Vec<_>> = entry.parts.iter().map(|anchor| anchor.resolve(pair)).collect();
        let mut parts = parts?.into_iter();
        let specifiers = parts.next()?;
        let declarator = parts.next()?;
        let body = parts.next_back()?;
        Some(FunctionRegion { specifiers, declarator, body, declarations: parts.collect(),
            value: entry.first.take().unwrap_or_else(function_shell) })
    })
}

/// Keep the walk_statement wrapper around this handoff: it owns timing and
/// source-span stamping. The pair is Rule::statement, not its unwrapped child.
pub(crate) fn take_control_statement(pair: &Pair<'_, Rule>) -> Option<OwnedStatement> {
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| cache.borrow_mut().as_mut()?
        .get_mut(&(source.as_ptr() as usize, source.len()))?.take_control(pair))
}

/// Resolve the entire ordered body before spending any owned values. A stale
/// capture graph falls back atomically, rather than consuming a partial body.
pub(crate) fn take_body<'a>(pair: &Pair<'a, Rule>) -> Option<Vec<BodyItem<'a>>> {
    let span = pair.as_span();
    let source = span.get_input();
    CACHE.with(|cache| cache.borrow_mut().as_mut()?
        .get_mut(&(source.as_ptr() as usize, source.len()))?.take_body(pair))
}

/// Read an initializer without spending the ordered walk's first owned value.
/// The callback holds no TLS borrow, and reparse invalidation wins over return.
pub(crate) fn with_initializer_list<T>(
    pair: &Pair<'_, Rule>, read: impl FnOnce(&Expression) -> T,
) -> Option<T> {
    let lease = lease_initializer_list(pair)?;
    Some(read(lease.as_ref()))
}

/// May be retained through source-order definition selection without a TLS
/// borrow. Only immutable AST access is exposed; drop restores the owned value.
pub(crate) struct InitializerLease {
    value: Option<Expression>,
    input: Input,
    span: (usize, usize),
    generation: Rc<()>,
}

impl AsRef<Expression> for InitializerLease {
    fn as_ref(&self) -> &Expression { self.value.as_ref().unwrap() }
}

impl Drop for InitializerLease {
    fn drop(&mut self) {
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let Some(values) = cache.as_mut().and_then(|cache| cache.get_mut(&self.input))
                else { return; };
            if !Rc::ptr_eq(&values.generation, &self.generation) { return; }
            let Some(entry) = values.entries.get_mut(&self.span) else { return; };
            if entry.first.is_none() { entry.first = self.value.take(); }
        });
    }
}

pub(crate) fn lease_initializer_list(pair: &Pair<'_, Rule>) -> Option<InitializerLease> {
    let span = pair.as_span();
    let source = span.get_input();
    let input = (source.as_ptr() as usize, source.len());
    CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let values = cache.as_mut()?.get_mut(&input)?;
        Some(InitializerLease { value: Some(values.take(pair)?), input,
            span: (span.start(), span.end()), generation: values.generation.clone() })
    })
}

fn split_array(node: Node, undo: Undo) -> Vec<Node> {
    let Node::Ready { value, recipe: Recipe::Array(recipes), .. } = node
        else { unreachable!("direct array inverse"); };
    let ExprKind::Array(values) = value.kind else { unreachable!("direct array AST"); };
    #[cfg(test)]
    if let Undo::ReferenceArray(origins) = undo {
        return origins.into_iter().zip(values).zip(recipes).map(|((origin, element), recipe)|
            Node::Ready { origin, value: element.value, recipe }).collect();
    }
    let Undo::Array(anchors) = undo else { unreachable!("direct array undo"); };
    let mut anchors = anchors.into_iter().peekable();
    values.into_iter().zip(recipes).enumerate().map(|(index, (element, recipe))| {
        let anchor = if anchors.peek().is_some_and(|(child, _)| *child == index) {
            Some(anchors.next().unwrap().1)
        } else { None };
        Node::Ready { origin: Origin { anchor }, value: element.value, recipe }
    }).collect()
}

fn unwrap_prefix(node: Node, undo: Undo) -> Node {
    let (Node::Ready { value, recipe: Recipe::Prefix { op, operand }, .. }, Undo::Prefix(origin)) =
        (node, undo) else { unreachable!("direct numeric prefix inverse"); };
    let value = if op == "+" { value } else {
        let ExprKind::Unary { expr, .. } = value.kind else { unreachable!("direct prefix AST"); };
        *expr
    };
    Node::Ready { origin, value, recipe: *operand }
}

fn unwrap_primary(node: Node, undo: Undo) -> Node {
    let (Node::Primary { value, recipe, .. }, Undo::Primary(origin)) = (node, undo)
        else { unreachable!("direct primary inverse"); };
    Node::Ready { origin, value, recipe }
}

fn syntax_child(node: Node) -> (Expression, Rc<SyntaxRecipe>, SyntaxOrigin) {
    match node {
        Node::Primary { span, value, recipe } =>
            (value, Rc::new(SyntaxRecipe::Primary { span, recipe }), SyntaxOrigin::Primary),
        Node::Syntax { value, recipe } => (value, recipe, SyntaxOrigin::Syntax),
        _ => unreachable!("direct syntax operand"),
    }
}

fn split_syntax(node: Node, undo: Undo) -> Vec<Node> {
    let (Node::Syntax { value, recipe }, Undo::Syntax(origins)) = (node, undo)
        else { unreachable!("direct syntax inverse"); };
    let recipe = Rc::try_unwrap(recipe).unwrap_or_else(|_| unreachable!("unpublished syntax owner"));
    let (values, recipes) = match recipe {
        SyntaxRecipe::Initializer { operand, .. } | SyntaxRecipe::Grouped { operand, .. } =>
            (vec![value], vec![operand]),
        SyntaxRecipe::Unary { op, operand, .. } => {
            let value = if op.is_some_and(|op| op != "+") {
                let ExprKind::Unary { expr, .. } = value.kind else { unreachable!(); };
                *expr
            } else { value };
            (vec![value], vec![operand])
        }
        SyntaxRecipe::Sequence { operands, .. } => {
            let values = if operands.len() == 1 { vec![value] } else {
                let ExprKind::Sequence(values) = value.kind else { unreachable!(); };
                values
            };
            (values, operands)
        }
        SyntaxRecipe::Primary { .. } => unreachable!("primary inverse has separate metadata"),
    };
    values.into_iter().zip(recipes).zip(origins).map(|((value, recipe), origin)| match origin {
        SyntaxOrigin::Syntax => Node::Syntax { value, recipe },
        SyntaxOrigin::Primary => {
            let recipe = Rc::try_unwrap(recipe).unwrap_or_else(|_| unreachable!("unpublished primary owner"));
            let SyntaxRecipe::Primary { span, recipe } = recipe else { unreachable!(); };
            Node::Primary { span, value, recipe }
        }
    }).collect()
}

fn unwrap_body_item(node: Node, undo: Undo) -> Node {
    let (Node::BodyItem(BodyReceipt { control: Some(control), .. }), Undo::BodyItem) = (node, undo)
        else { unreachable!("direct statement inverse"); };
    Node::Control(control)
}

fn unwrap_plain_name(node: Node, undo: Undo) -> Node {
    let (Node::PlainName { name, .. }, Undo::PlainName) = (node, undo)
        else { unreachable!("direct declarator name inverse"); };
    Node::Label(name)
}

fn split_declarator(node: Node, undo: Undo) -> Vec<Node> {
    let (Node::Declarator { value, initializer, .. }, Undo::Declarator(span)) = (node, undo)
        else { unreachable!("direct declarator inverse"); };
    let mut children = vec![Node::PlainName { span, name: value.syntax.declarator }];
    if let Some(recipe) = initializer {
        children.push(Node::Syntax { value: value.declaration.init.unwrap(), recipe });
    }
    children
}

fn split_parameter(node: Node, undo: Undo) -> Vec<Node> {
    let (Node::Parameter { recipe, .. }, Undo::Parameter(plain_span)) = (node, undo)
        else { unreachable!("direct parameter inverse"); };
    vec![Node::ParameterSpecifiers { span: recipe.specifier_span, capture: Some(recipe.specifiers) },
        Node::PlainName { span: plain_span, name: recipe.name }]
}

fn binary_child(node: Node) -> (Expression, Rc<BinaryRecipe>, BinaryOrigin) {
    match node {
        Node::Syntax { value, recipe } => {
            let SyntaxRecipe::Unary { capture, .. } = &*recipe else { unreachable!("original unary operand"); };
            (value, Rc::new(BinaryRecipe::Atom { capture: *capture, syntax: recipe }), BinaryOrigin::Syntax)
        }
        Node::Binary { span, value, recipe } => (value, recipe, BinaryOrigin::Binary(span)),
        _ => unreachable!("direct binary operand"),
    }
}

fn split_binary(node: Node, undo: Undo) -> Vec<Node> {
    let (Node::Binary { value, recipe, .. }, Undo::Binary(origins)) = (node, undo)
        else { unreachable!("direct binary inverse"); };
    let ExprKind::Binary { left, right, .. } = value.kind else { unreachable!(); };
    let recipe = Rc::try_unwrap(recipe).unwrap_or_else(|_| unreachable!("unpublished binary owner"));
    let BinaryRecipe::Infix { left: lrecipe, right: rrecipe, .. } = recipe else { unreachable!(); };
    [(*left, lrecipe), (*right, rrecipe)].into_iter().zip(origins).map(|((value, recipe), origin)| {
        match origin {
            BinaryOrigin::Binary(span) => Node::Binary { span, value, recipe },
            BinaryOrigin::Syntax => {
                let recipe = Rc::try_unwrap(recipe).unwrap_or_else(|_| unreachable!("unpublished binary atom"));
                let BinaryRecipe::Atom { syntax: recipe, .. } = recipe else { unreachable!(); };
                Node::Syntax { value, recipe }
            }
        }
    }).collect()
}

fn unwrap_binary_span(node: Node, undo: Undo) -> Node {
    let (Node::Binary { value, recipe, .. }, Undo::BinarySpan(span)) = (node, undo)
        else { unreachable!("direct binary region inverse"); };
    Node::Binary { span, value, recipe }
}

fn unwrap_control_label(node: Node, undo: Undo) -> Node {
    let (Node::Control(control), Undo::Label) = (node, undo)
        else { unreachable!("direct control label inverse"); };
    let span = match control.recipe {
        ControlRecipe::Continue(Some(span)) | ControlRecipe::Goto(span) => span,
        _ => unreachable!("direct labeled control recipe"),
    };
    Node::Label(span)
}

fn split_body(node: Node) -> Vec<Node> {
    let Node::Body { items, .. } = node else { unreachable!("direct body inverse"); };
    items.into_iter().map(Node::BodyItem).collect()
}

struct Cached {
    first: Option<Expression>,
    recipe: Recipe,
}

/// A committed, source-identity-bound handoff. No AST/scopes escape speculation.
pub(crate) struct Initializers {
    input: (usize, usize),
    entries: HashMap<(usize, usize), Cached>,
    primaries: HashMap<(usize, usize), Cached>,
    syntax: HashMap<(Rule, usize, usize), CachedSyntax>,
    declarators: HashMap<(usize, usize), CachedDeclarator>,
    binaries: HashMap<(usize, usize), CachedBinary>,
    parameters: HashMap<(usize, usize), CachedParameter>,
    ifs: HashMap<(usize, usize), CachedIf>,
    loops: HashMap<(Rule, usize, usize), CachedLoop>,
    fors: HashMap<(usize, usize), CachedFor>,
    switches: HashMap<(usize, usize), CachedSwitch>,
    returns: HashMap<(usize, usize), CachedReturn>,
    calls: HashMap<(usize, usize), CachedCall>,
    indices: HashMap<(usize, usize), CachedPlace>,
    members: HashMap<(Rule, usize, usize), CachedPlace>,
    functions: HashMap<(usize, usize), CachedFunction>,
    statements: HashMap<(usize, usize), CachedStatement>,
    bodies: HashMap<(usize, usize), Vec<CaptureHandle>>,
    generation: Rc<()>,
}

impl Initializers {
    fn append(&mut self, additions: Self) {
        debug_assert_eq!(self.input, additions.input);
        // Existing leases and repeated reads belong to the same generation.
        // Replayed/overlapping unit receipts must not spend or replace them.
        for (span, value) in additions.entries { self.entries.entry(span).or_insert(value); }
        for (span, value) in additions.primaries { self.primaries.entry(span).or_insert(value); }
        for (span, value) in additions.syntax { self.syntax.entry(span).or_insert(value); }
        for (span, value) in additions.declarators { self.declarators.entry(span).or_insert(value); }
        for (span, value) in additions.binaries { self.binaries.entry(span).or_insert(value); }
        for (span, value) in additions.parameters { self.parameters.entry(span).or_insert(value); }
        for (span, value) in additions.ifs { self.ifs.entry(span).or_insert(value); }
        for (span, value) in additions.loops { self.loops.entry(span).or_insert(value); }
        for (span, value) in additions.fors { self.fors.entry(span).or_insert(value); }
        for (span, value) in additions.switches { self.switches.entry(span).or_insert(value); }
        for (span, value) in additions.returns { self.returns.entry(span).or_insert(value); }
        for (span, value) in additions.calls { self.calls.entry(span).or_insert(value); }
        for (span, value) in additions.indices { self.indices.entry(span).or_insert(value); }
        for (span, value) in additions.members { self.members.entry(span).or_insert(value); }
        for (span, value) in additions.functions { self.functions.entry(span).or_insert(value); }
        for (span, value) in additions.statements { self.statements.entry(span).or_insert(value); }
        for (span, value) in additions.bodies { self.bodies.entry(span).or_insert(value); }
    }

    fn store_control(&mut self, receipt: BodyReceipt) {
        if let Some(control) = receipt.control {
            let span = receipt.span;
            self.statements.insert((span.start, span.end), CachedStatement {
                first: Some(control.value), recipe: control.recipe,
            });
        }
    }

    fn store_syntax(&mut self, recipe: Rc<SyntaxRecipe>, first: Option<Expression>) {
        let (rule, span) = recipe.rule_span();
        match &*recipe {
            SyntaxRecipe::Primary { .. } => {}
            SyntaxRecipe::Unary { operand, .. } | SyntaxRecipe::Initializer { operand, .. }
                | SyntaxRecipe::Grouped { operand, .. } =>
                self.store_syntax(operand.clone(), None),
            SyntaxRecipe::Sequence { operands, .. } => {
                for operand in operands { self.store_syntax(operand.clone(), None); }
            }
        }
        self.syntax.entry((rule, span.start, span.end)).or_insert(CachedSyntax { first, recipe });
    }

    fn take_syntax(&mut self, pair: &Pair<'_, Rule>) -> Option<Expression> {
        let span = pair.as_span();
        let source = span.get_input();
        if self.input != (source.as_ptr() as usize, source.len()) { return None; }
        let entry = self.syntax.get_mut(&(pair.as_rule(), span.start(), span.end()))?;
        Some(entry.first.take().unwrap_or_else(|| entry.recipe.replay(source)))
    }

    fn store_binary_syntax(&mut self, recipe: &BinaryRecipe) {
        let mut pending = vec![recipe];
        while let Some(recipe) = pending.pop() {
            match recipe {
                BinaryRecipe::Atom { syntax, .. } => self.store_syntax(syntax.clone(), None),
                BinaryRecipe::Infix { left, right, .. } => { pending.push(right); pending.push(left); }
            }
        }
    }

    fn take_control(&mut self, pair: &Pair<'_, Rule>) -> Option<OwnedStatement> {
        if pair.as_rule() != Rule::statement { return None; }
        let span = pair.as_span();
        let source = span.get_input();
        if self.input != (source.as_ptr() as usize, source.len()) { return None; }
        let entry = self.statements.get_mut(&(span.start(), span.end()))?;
        Some(entry.first.take().unwrap_or_else(|| entry.recipe.replay(source)))
    }

    fn take_body<'a>(&mut self, pair: &Pair<'a, Rule>) -> Option<Vec<BodyItem<'a>>> {
        if pair.as_rule() != Rule::compound_statement { return None; }
        let span = pair.as_span();
        let source = span.get_input();
        if self.input != (source.as_ptr() as usize, source.len()) { return None; }
        let handles = self.bodies.get(&(span.start(), span.end()))?;
        let anchors: Vec<_> = handles.iter().map(|handle| pair.resolve_capture(*handle))
            .collect::<Option<_>>()?;
        if anchors.iter().any(|anchor| anchor.as_rule() != Rule::statement) { return None; }
        Some(anchors.into_iter().map(|pair| match self.take_control(&pair) {
            Some(value) => BodyItem::Owned { pair, value },
            None => BodyItem::Contextual(pair),
        }).collect())
    }

    pub(crate) fn take(&mut self, pair: &Pair<'_, Rule>) -> Option<Expression> {
        if pair.as_rule() != Rule::initializer_list { return None; }
        let span = pair.as_span();
        let source = span.get_input();
        if self.input != (source.as_ptr() as usize, source.len()) { return None; }
        let entry = self.entries.get_mut(&(span.start(), span.end()))?;
        Some(entry.first.take().unwrap_or_else(|| entry.recipe.replay(source)))
    }

    pub(crate) fn len(&self) -> usize { self.entries.len() }
}

/// Use in the same hybrid match that produces the walker's initializer anchors.
/// `new` retains captures until the consumer audit permits boundary omission.
pub(crate) struct Direct<'a> {
    source: &'a str,
    arena: EditArena<Node, Undo, ()>,
    depth: usize,
    literal_active: bool,
    state_undo: Vec<(usize, StateUndo)>,
    frames: Vec<Frame>,
    committed: bool,
    elide: bool,
    admitted_end: Option<(usize, usize)>,
    #[cfg(test)]
    reference_construction: bool,
}

impl<'a> Direct<'a> {
    pub(crate) fn new(source: &'a str) -> Self {
        Self { source, arena: Default::default(), depth: 0, literal_active: false,
            state_undo: Vec::new(), frames: Vec::new(),
            committed: false, elide: false, admitted_end: None,
            #[cfg(test)]
            reference_construction: REFERENCE_CONSTRUCTION.with(|slot| slot.get()),
        }
    }

    /// Requires grammar-owned initializer_list capture-boundary certification
    /// and consumers that take the complete owned AST before visiting children.
    pub(crate) fn with_capture_elision(source: &'a str) -> Self {
        Self { elide: true, ..Self::new(source) }
    }

    pub(crate) fn into_initializers(self) -> Initializers {
        let mut result = Initializers {
            input: (self.source.as_ptr() as usize, self.source.len()), entries: HashMap::new(),
            primaries: HashMap::new(), syntax: HashMap::new(), declarators: HashMap::new(), binaries: HashMap::new(),
            statements: HashMap::new(), bodies: HashMap::new(), parameters: HashMap::new(),
            ifs: HashMap::new(),
            loops: HashMap::new(),
            fors: HashMap::new(),
            switches: HashMap::new(),
            returns: HashMap::new(),
            calls: HashMap::new(),
            indices: HashMap::new(), members: HashMap::new(),
            functions: HashMap::new(),
            generation: Rc::new(()),
        };
        if self.committed {
            for node in self.arena.into_live_nodes() {
                match node {
                    Node::Ready { origin: Origin { anchor: Some(span), .. }, value, recipe } => {
                        result.entries.insert((span.start, span.end), Cached { first: Some(value), recipe });
                    }
                    Node::Primary { span, value, recipe } => {
                        result.primaries.insert((span.start, span.end), Cached { first: Some(value), recipe });
                    }
                    Node::Syntax { value, recipe } => result.store_syntax(recipe, Some(value)),
                    Node::Declarator { span, value, initializer } => {
                        if let Some(recipe) = &initializer { result.store_syntax(recipe.clone(), None); }
                        let syntax = value.syntax;
                        result.declarators.insert((span.start, span.end), CachedDeclarator {
                            first: Some(value), syntax, initializer,
                        });
                    }
                    Node::Binary { span, value, recipe } => {
                        result.store_binary_syntax(&recipe);
                        if let Some(span) = span {
                            result.binaries.insert((span.start, span.end), CachedBinary { first: Some(value), recipe });
                        }
                    }
                    Node::Parameter { span, value, recipe } => {
                        result.parameters.insert((span.start, span.end), CachedParameter { first: Some(value), recipe });
                    }
                    Node::IfRegion { span, value, recipe } => {
                        result.ifs.insert((span.start, span.end), CachedIf { first: Some(value), recipe });
                    }
                    Node::LoopRegion { span, value, recipe } => {
                        let rule = if recipe.do_while { Rule::do_while_statement } else { Rule::while_statement };
                        result.loops.insert((rule, span.start, span.end), CachedLoop { first: Some(value), recipe });
                    }
                    Node::ForRegion { span, value, recipe } => {
                        result.fors.insert((span.start, span.end), CachedFor { first: Some(value), recipe });
                    }
                    Node::SwitchRegion { span, value, recipe } => {
                        result.switches.insert((span.start, span.end), CachedSwitch { first: Some(value), recipe });
                    }
                    Node::ReturnRegion { span, value, expression } => {
                        result.returns.insert((span.start, span.end), CachedReturn { first: Some(value), expression });
                    }
                    Node::CallRegion { span, value, arguments } => {
                        result.calls.insert((span.start, span.end), CachedCall { first: Some(value), arguments });
                    }
                    Node::IndexRegion { span, value, anchor } => {
                        result.indices.insert((span.start, span.end), CachedPlace { first: Some(value), anchor });
                    }
                    Node::MemberRegion { rule, span, value, anchor } => {
                        result.members.insert((rule, span.start, span.end), CachedPlace { first: Some(value), anchor });
                    }
                    Node::FunctionRegion { span, value, parts } => {
                        result.functions.insert((span.start, span.end), CachedFunction { first: Some(value), parts });
                    }
                    Node::BodyItem(receipt) => result.store_control(receipt),
                    Node::Body { span, items } => {
                        result.bodies.insert((span.start, span.end), items.iter().map(|item| item.capture).collect());
                        for item in items { result.store_control(item); }
                    }
                    _ => {}
                }
            }
        }
        result
    }

    fn push(&mut self, node: Node) {
        let id = self.arena.alloc(node);
        self.arena.push(id);
    }

    fn set_state(&mut self, depth: usize, literal_active: bool) {
        self.state_undo.push((self.arena.checkpoint(),
            StateUndo::State { depth: self.depth, literal_active: self.literal_active }));
        self.depth = depth;
        self.literal_active = literal_active;
    }

    fn literal(rule: usize) -> Option<Rule> {
        if rule == Rule::int_literal as usize { Some(Rule::int_literal) }
        else if rule == Rule::float_literal as usize { Some(Rule::float_literal) }
        else if rule == Rule::char_literal as usize { Some(Rule::char_literal) }
        else if rule == Rule::utf8_char_literal as usize { Some(Rule::utf8_char_literal) }
        else if rule == Rule::wide_char_literal as usize { Some(Rule::wide_char_literal) }
        else if rule == Rule::string_literal as usize { Some(Rule::string_literal) }
        else if rule == Rule::wide_string_literal as usize { Some(Rule::wide_string_literal) }
        else if rule == Rule::bool_literal as usize { Some(Rule::bool_literal) }
        else { None }
    }

    fn wrapper(rule: usize) -> bool {
        rule == Rule::initializer as usize || rule == Rule::designated_init as usize
            || rule == Rule::assignment_expression as usize || rule == Rule::conditional_expression as usize
            || rule == Rule::binary_expression as usize || rule == Rule::unary_expression as usize
            || rule == Rule::postfix_expression as usize || rule == Rule::primary_expression as usize
            || rule == Rule::literal as usize || rule == Rule::expression as usize
    }

    fn current_frame(&self) -> Option<(usize, usize)> {
        #[cfg(test)]
        if self.reference_construction() {
            return self.arena.values().iter().rev().find_map(|id| match self.arena.get(*id) {
                Node::Frame { rule, start } => Some((*rule, *start)), _ => None,
            });
        }
        self.frames.last().map(|frame| (frame.rule, frame.start))
    }

    fn body_rule(rule: usize) -> bool {
        rule == Rule::compound_statement as usize || rule == Rule::statement as usize
            || rule == Rule::break_statement as usize || rule == Rule::continue_statement as usize
            || rule == Rule::goto_statement as usize || rule == Rule::empty_statement as usize
            || rule == Rule::return_statement as usize
            || rule == Rule::if_statement as usize || rule == Rule::else_clause as usize
            || rule == Rule::while_statement as usize || rule == Rule::do_while_statement as usize
            || [Rule::for_statement, Rule::for_init, Rule::for_cond, Rule::for_update]
                .iter().any(|candidate| rule == *candidate as usize)
            || rule == Rule::switch_statement as usize
            || rule == Rule::function_definition as usize
    }

    fn syntax_rule(rule: usize) -> bool {
        [Rule::expression, Rule::assignment_expression, Rule::conditional_expression,
            Rule::binary_expression, Rule::binary_op, Rule::unary_expression,
            Rule::prefix_op, Rule::postfix_expression, Rule::assignment_op,
            Rule::call_suffix, Rule::index_suffix, Rule::member_suffix,
            Rule::argument_list,
            Rule::arrow_suffix, Rule::inc_dec_suffix, Rule::cast_expression,
            Rule::sizeof_expression, Rule::alignof_expression, Rule::offsetof_expression,
            Rule::va_arg_expression, Rule::generic_expression,
            Rule::statement_expression, Rule::compound_literal,
        ].iter().any(|candidate| rule == *candidate as usize)
    }

    fn finish_with_capture(&mut self, rule: usize, span: Span, source: &str,
        capture: Option<CaptureHandle>)
    {
        if self.current_frame() != Some((rule, span.start)) { return; }
        self.finish_node(rule, span, source, capture);
        if self.current_frame().is_some_and(|(parent, _)| parent == Rule::function_definition as usize)
            && [Rule::declaration_specifiers, Rule::declarator, Rule::normal_declaration,
                Rule::compound_statement].iter().any(|candidate| rule == *candidate as usize)
            && let Some(capture) = capture
        {
            self.arena.pop().expect("finished function syntax part");
            let rule = if rule == Rule::declaration_specifiers as usize { Rule::declaration_specifiers }
                else if rule == Rule::declarator as usize { Rule::declarator }
                else if rule == Rule::normal_declaration as usize { Rule::normal_declaration }
                else { Rule::compound_statement };
            self.push(Node::ContextAnchor(ContextAnchor { rule, span, capture }));
        }
        if (rule == Rule::expression as usize
                && self.current_frame().is_some_and(|(parent, _)| parent == Rule::index_suffix as usize)
            || rule == Rule::ident_name as usize
                && self.current_frame().is_some_and(|(parent, _)|
                    parent == Rule::member_suffix as usize || parent == Rule::arrow_suffix as usize))
            && let Some(capture) = capture
        {
            self.arena.pop().expect("finished contextual place operand");
            let rule = if rule == Rule::ident_name as usize { Rule::ident_name } else { Rule::expression };
            self.push(Node::ContextAnchor(ContextAnchor { rule, span, capture }));
        }
        if rule == Rule::assignment_expression as usize
            && self.current_frame().is_some_and(|(parent, _)| parent == Rule::argument_list as usize)
            && let Some(capture) = capture
        {
            self.arena.pop().expect("finished contextual argument");
            self.push(Node::ContextAnchor(ContextAnchor { rule: Rule::assignment_expression, span, capture }));
        }
        if [Rule::for_init, Rule::for_cond, Rule::for_update].iter().any(|candidate| rule == *candidate as usize)
            && self.current_frame().is_some_and(|(parent, _)| parent == Rule::for_statement as usize)
            && let Some(capture) = capture
        {
            self.arena.pop().expect("finished for clause");
            let rule = if rule == Rule::for_init as usize { Rule::for_init }
                else if rule == Rule::for_cond as usize { Rule::for_cond } else { Rule::for_update };
            self.push(Node::ContextAnchor(ContextAnchor { rule, span, capture }));
        }
        if rule == Rule::expression as usize
            && self.current_frame().is_some_and(|(parent, _)|
                [Rule::if_statement, Rule::while_statement, Rule::do_while_statement, Rule::switch_statement,
                    Rule::return_statement]
                    .iter().any(|rule| parent == *rule as usize))
            && let Some(capture) = capture
        {
            // Keep the expression node arena-owned for its independent handoff;
            // only its structural stack slot becomes the contextual anchor.
            self.arena.pop().expect("finished contextual condition");
            self.push(Node::ContextAnchor(ContextAnchor { rule: Rule::expression, span, capture }));
        }
        if Self::literal(rule).is_some() { self.set_state(self.depth, false); }
        if rule == Rule::initializer_list as usize { self.set_state(self.depth - 1, false); }
    }

    #[inline(always)]
    fn reference_construction(&self) -> bool {
        #[cfg(test)]
        { self.reference_construction }
        #[cfg(not(test))]
        { false }
    }

    #[cfg(test)]
    fn finish_reference_array(&mut self, children: Vec<grameo::arena::Id>, span: Span) {
        let id = self.arena.list_with_metadata(children, |children| {
            let mut values = Vec::with_capacity(children.len());
            let mut recipes = Vec::with_capacity(children.len());
            let mut origins = Vec::with_capacity(children.len());
            for child in children {
                let Node::Ready { origin, value, recipe } = child else { unreachable!(); };
                origins.push(origin);
                values.push(value);
                recipes.push(recipe);
            }
            (Node::Ready { origin: Origin { anchor: Some(span) }, value: array(values),
                recipe: Recipe::Array(recipes) }, Undo::ReferenceArray(origins))
        }, split_array);
        self.arena.push(id);
    }

    fn finish_node(&mut self, rule: usize, span: Span, source: &str, capture: Option<CaptureHandle>) {
        let base = if self.reference_construction() { None } else {
            let frame = self.frames.pop().expect("direct semantic frame");
            assert_eq!((frame.rule, frame.start), (rule, span.start));
            self.state_undo.push((self.arena.checkpoint(), StateUndo::PushFrame(frame)));
            // The common wrapper changes neither the AST nor its stack slot.
            if Self::wrapper(rule) && !(self.depth == 0 && rule == Rule::primary_expression as usize)
                && self.arena.values().len() == frame.base + 1
                && matches!(self.arena.get(self.arena.values()[frame.base]), Node::Ready { .. })
            { return; }
            Some(frame.base)
        };
        let mut children = Vec::new();
        while base.is_none_or(|base| self.arena.values().len() > base) {
            let Some(id) = self.arena.pop() else { break; };
            #[cfg(test)]
            if self.reference_construction()
                && matches!(self.arena.get(id), Node::Frame { rule: owner, start }
                    if *owner == rule && *start == span.start) { break; }
            children.push(id);
        }
        children.reverse();
        if rule == Rule::function_definition as usize && children.len() >= 3 {
            let parts: Option<Vec<_>> = children.iter().map(|id| match self.arena.get(*id) {
                Node::ContextAnchor(anchor) => Some(*anchor), _ => None,
            }).collect();
            if let Some(parts) = parts
                && parts[0].rule == Rule::declaration_specifiers && parts[1].rule == Rule::declarator
                && parts.last().is_some_and(|part| part.rule == Rule::compound_statement)
                && parts[2..parts.len()-1].iter().all(|part| part.rule == Rule::normal_declaration)
            {
                self.push(Node::FunctionRegion { span, value: function_shell(), parts });
                return;
            }
        }
        if children.len() == 1 && let Node::ContextAnchor(anchor) = self.arena.get(children[0]) {
            let anchor = *anchor;
            if rule == Rule::return_statement as usize && anchor.rule == Rule::expression {
                self.push(Node::ReturnRegion { span, value: return_shell(), expression: anchor });
                return;
            }
            if rule == Rule::index_suffix as usize && anchor.rule == Rule::expression {
                self.push(Node::IndexRegion { span, value: index_shell(), anchor });
                return;
            }
            if (rule == Rule::member_suffix as usize || rule == Rule::arrow_suffix as usize)
                && anchor.rule == Rule::ident_name
            {
                let rule = if rule == Rule::member_suffix as usize { Rule::member_suffix } else { Rule::arrow_suffix };
                self.push(Node::MemberRegion { rule, span,
                    value: member_shell(&source[anchor.span.start..anchor.span.end]), anchor });
                return;
            }
        }
        if rule == Rule::argument_list as usize && !children.is_empty()
            && children.iter().all(|id| matches!(self.arena.get(*id),
                Node::ContextAnchor(anchor) if anchor.rule == Rule::assignment_expression))
        {
            let arguments = children.iter().map(|id| match self.arena.get(*id) {
                Node::ContextAnchor(anchor) => *anchor, _ => unreachable!(),
            }).collect();
            self.push(Node::CallArguments(arguments));
            return;
        }
        if rule == Rule::call_suffix as usize && (children.is_empty()
            || children.len() == 1 && matches!(self.arena.get(children[0]), Node::CallArguments(_)))
        {
            let arguments = if children.is_empty() { Vec::new() } else {
                let Node::CallArguments(arguments) = self.arena.get(children[0]) else { unreachable!(); };
                arguments.clone()
            };
            self.push(Node::CallRegion { span, value: call_shell(), arguments });
            return;
        }
        if rule == Rule::switch_statement as usize && children.len() == 2
            && let Node::ContextAnchor(subject) = self.arena.get(children[0])
            && subject.rule == Rule::expression
            && let Node::BodyItem(body) = self.arena.get(children[1])
        {
            let recipe = SwitchRecipe { subject: *subject,
                body: ContextAnchor { rule: Rule::statement, span: body.span, capture: body.capture } };
            self.push(Node::SwitchRegion { span, value: switch_shell(), recipe });
            return;
        }
        if rule == Rule::for_statement as usize && (children.len() == 3 || children.len() == 4)
            && let Node::ContextAnchor(init) = self.arena.get(children[0])
            && init.rule == Rule::for_init
            && let Node::ContextAnchor(condition) = self.arena.get(children[1])
            && condition.rule == Rule::for_cond
            && let Node::BodyItem(body) = self.arena.get(*children.last().unwrap())
        {
            let update = if children.len() == 4 {
                match self.arena.get(children[2]) {
                    Node::ContextAnchor(anchor) if anchor.rule == Rule::for_update => Some(*anchor),
                    _ => { self.push(Node::Unsupported); return; }
                }
            } else { None };
            let recipe = ForRecipe { init: *init, condition: *condition, update,
                body: ContextAnchor { rule: Rule::statement, span: body.span, capture: body.capture } };
            self.push(Node::ForRegion { span, value: for_shell(), recipe });
            return;
        }
        if (rule == Rule::while_statement as usize || rule == Rule::do_while_statement as usize)
            && children.len() == 2
        {
            let do_while = rule == Rule::do_while_statement as usize;
            let (condition, body) = if do_while { (children[1], children[0]) }
                else { (children[0], children[1]) };
            if let Node::ContextAnchor(condition) = self.arena.get(condition)
                && condition.rule == Rule::expression
                && let Node::BodyItem(body) = self.arena.get(body)
            {
                let recipe = LoopRecipe { condition: *condition,
                    body: ContextAnchor { rule: Rule::statement, span: body.span, capture: body.capture },
                    do_while };
                self.push(Node::LoopRegion { span, value: recipe.shell(), recipe });
                return;
            }
        }
        if rule == Rule::else_clause as usize && children.len() == 1
            && let Node::BodyItem(item) = self.arena.get(children[0])
        {
            self.push(Node::ContextAnchor(ContextAnchor {
                rule: Rule::statement, span: item.span, capture: item.capture,
            }));
            return;
        }
        if rule == Rule::if_statement as usize && (children.len() == 2 || children.len() == 3)
            && let Node::ContextAnchor(condition) = self.arena.get(children[0])
            && condition.rule == Rule::expression
            && let Node::BodyItem(then_branch) = self.arena.get(children[1])
        {
            let else_branch = if children.len() == 3 {
                match self.arena.get(children[2]) {
                    Node::ContextAnchor(anchor) if anchor.rule == Rule::statement => Some(*anchor),
                    _ => { self.push(Node::Unsupported); return; }
                }
            } else { None };
            let recipe = IfRecipe { condition: *condition,
                then_branch: ContextAnchor { rule: Rule::statement,
                    span: then_branch.span, capture: then_branch.capture }, else_branch };
            // Children remain live for independent syntax/body/control handoffs.
            // Stack edits, not destructive AST moves, own this wrapper's undo.
            self.push(Node::IfRegion { span, value: if_shell(), recipe });
            return;
        }
        if rule == Rule::ident_name as usize && children.is_empty() {
            if self.current_frame().is_some_and(|(parent, _)| parent == Rule::primary_expression as usize) {
                let recipe = Recipe::Ident(span);
                self.push(Node::Ready { origin: Origin { anchor: None }, value: recipe.replay(source), recipe });
            } else { self.push(Node::Label(span)); }
            return;
        }
        if rule == Rule::primary_expression as usize && self.depth == 0
            && children.len() == 1 && matches!(self.arena.get(children[0]), Node::Ready { .. })
        {
            let id = self.arena.unary_with_metadata(children[0], |node| {
                let Node::Ready { origin, value, recipe } = node else { unreachable!(); };
                (Node::Primary { span, value, recipe }, Undo::Primary(origin))
            }, unwrap_primary);
            self.arena.push(id);
            return;
        }
        if self.depth == 0 {
            let syntax = |node: &Node| matches!(node, Node::Primary { .. } | Node::Syntax { .. });
            if rule == Rule::declaration_specifiers as usize {
                // Complex specifier forms keep contextual metadata handling.
                // Their syntax guards leave children; primitive/alias spelling
                // is retained as a capture for the ordered type policy.
                self.push(Node::ParameterSpecifiers { span,
                    capture: if children.is_empty() { capture } else { None } });
                return;
            }
            if rule == Rule::parameter_decl as usize && children.len() == 2
                && let Node::ParameterSpecifiers { span: specs, capture: Some(handle) } = self.arena.get(children[0])
                && let Node::PlainName { span: plain, name } = self.arena.get(children[1])
                && specs.start == span.start && specs.end <= plain.start && plain.end <= span.end
                && source.get(specs.end..plain.start).is_some_and(super::owned::plain_declarator_trivia)
                && source.get(plain.end..span.end).is_some_and(super::owned::plain_declarator_trivia)
            {
                let recipe = ParameterRecipe { name: *name, specifiers: *handle, specifier_span: *specs,
                    declarator_span: Span { start: plain.start, end: span.end } };
                let plain_span = *plain;
                let id = self.arena.list_with_metadata(children, |_| {
                    let value = crate::walker::construct_plain_parameter(source[recipe.name.start..recipe.name.end].to_string());
                    (Node::Parameter { span, value, recipe }, Undo::Parameter(plain_span))
                }, split_parameter);
                self.arena.push(id);
                return;
            }
            if rule == Rule::binary_expression as usize && children.len() == 1
                && matches!(self.arena.get(children[0]), Node::Binary { .. })
            {
                let id = self.arena.unary_with_metadata(children[0], |node| {
                    let Node::Binary { span: previous, value, recipe } = node else { unreachable!(); };
                    (Node::Binary { span: Some(span), value, recipe }, Undo::BinarySpan(previous))
                }, unwrap_binary_span);
                self.arena.push(id);
                return;
            }
            if rule == Rule::primary_expression as usize && children.len() == 1
                && matches!(self.arena.get(children[0]), Node::Syntax { recipe, .. }
                    if recipe.rule_span().0 == Rule::expression)
            {
                let id = self.arena.list_with_metadata(children, |mut nodes| {
                    let (value, operand, origin) = syntax_child(nodes.pop().unwrap());
                    (Node::Syntax { value, recipe: Rc::new(SyntaxRecipe::Grouped { span, operand }) },
                        Undo::Syntax(vec![origin]))
                }, split_syntax);
                self.arena.push(id);
                return;
            }
            if rule == Rule::direct_declarator as usize && children.len() == 1
                && let Node::Label(name) = self.arena.get(children[0])
                && name.start == span.start && name.end <= span.end
                && source.get(name.end..span.end).is_some_and(super::owned::plain_declarator_trivia)
            {
                let name = *name;
                let id = self.arena.unary_with_metadata(children[0], |_| {
                    (Node::PlainName { span, name }, Undo::PlainName)
                }, unwrap_plain_name);
                self.arena.push(id);
                return;
            }
            if rule == Rule::declarator as usize && children.len() == 1
                && let Node::PlainName { span: plain, .. } = self.arena.get(children[0])
                && plain.start == span.start && plain.end <= span.end
                && source.get(plain.end..span.end).is_some_and(super::owned::plain_declarator_trivia)
            {
                self.arena.push(children[0]);
                return;
            }
            if rule == Rule::init_declarator as usize && matches!(children.len(), 1 | 2)
                && matches!(self.arena.get(children[0]), Node::PlainName { span: plain, .. } if plain.start == span.start)
                && (children.len() == 1 || matches!(self.arena.get(children[1]),
                    Node::Syntax { recipe, .. } if recipe.rule_span().0 == Rule::initializer))
            {
                let id = self.arena.list_with_metadata(children, |nodes| {
                    let mut nodes = nodes.into_iter();
                    let Node::PlainName { span: plain_span, name } = nodes.next().unwrap() else { unreachable!(); };
                    let (init, initializer) = match nodes.next() {
                        Some(Node::Syntax { value, recipe }) => (Some(value), Some(recipe)),
                        None => (None, None),
                        _ => unreachable!(),
                    };
                    let syntax = PlainDeclaratorSyntax { declarator: name,
                        initializer: initializer.as_ref().map(|recipe| recipe.rule_span().1) };
                    let value = construct_plain_declarator(source[name.start..name.end].to_string(), init, syntax);
                    (Node::Declarator { span, value, initializer }, Undo::Declarator(plain_span))
                }, split_declarator);
                self.arena.push(id);
                return;
            }
            if rule == Rule::initializer as usize && children.len() == 1
                && syntax(self.arena.get(children[0]))
            {
                let id = self.arena.list_with_metadata(children, |mut nodes| {
                    let (value, operand, origin) = syntax_child(nodes.pop().unwrap());
                    (Node::Syntax { value, recipe: Rc::new(SyntaxRecipe::Initializer { span, operand }) },
                        Undo::Syntax(vec![origin]))
                }, split_syntax);
                self.arena.push(id);
                return;
            }
            if rule == Rule::expression as usize && !children.is_empty()
                && children.iter().all(|id| syntax(self.arena.get(*id)))
            {
                let id = self.arena.list_with_metadata(children, |nodes| {
                    let mut values = Vec::with_capacity(nodes.len());
                    let mut operands = Vec::with_capacity(nodes.len());
                    let mut origins = Vec::with_capacity(nodes.len());
                    for node in nodes {
                        let (value, recipe, origin) = syntax_child(node);
                        values.push(value); operands.push(recipe); origins.push(origin);
                    }
                    let value = if values.len() == 1 { values.pop().unwrap() }
                        else { Expression::new(ExprKind::Sequence(values)) };
                    (Node::Syntax { value, recipe: Rc::new(SyntaxRecipe::Sequence { span, operands }) }, Undo::Syntax(origins))
                }, split_syntax);
                self.arena.push(id);
                return;
            }
            if rule == Rule::unary_expression as usize {
                let operand = if children.len() == 1 && syntax(self.arena.get(children[0])) {
                    Some((children[0], None))
                } else if children.len() == 2 && syntax(self.arena.get(children[1]))
                    && let Node::Prefix(op) = self.arena.get(children[0]) {
                    Some((children[1], Some(*op)))
                } else { None };
                if let Some((operand, op)) = operand {
                    let id = self.arena.list_with_metadata(vec![operand], |mut nodes| {
                        let (value, operand, origin) = syntax_child(nodes.pop().unwrap());
                        let value = if let Some(op) = op { crate::walker::construct_numeric_prefix(op, value) } else { value };
                        (Node::Syntax { value, recipe: Rc::new(SyntaxRecipe::Unary { span, op, operand, capture }) }, Undo::Syntax(vec![origin]))
                    }, split_syntax);
                    self.arena.push(id);
                    return;
                }
            }
            if [Rule::assignment_expression, Rule::conditional_expression,
                Rule::binary_expression, Rule::postfix_expression].iter().any(|candidate| rule == *candidate as usize)
                && children.len() == 1 && syntax(self.arena.get(children[0]))
            {
                self.arena.push(children[0]);
                return;
            }
        }
        let control = if children.is_empty() {
            if rule == Rule::break_statement as usize { Some(ControlRecipe::Break) }
            else if rule == Rule::continue_statement as usize { Some(ControlRecipe::Continue(None)) }
            else if rule == Rule::return_statement as usize { Some(ControlRecipe::ReturnVoid) }
            else if rule == Rule::empty_statement as usize { Some(ControlRecipe::Empty) }
            else { None }
        } else { None };
        if let Some(recipe) = control {
            self.push(Node::Control(Control { value: recipe.replay(source), recipe }));
            return;
        }
        if (rule == Rule::goto_statement as usize || rule == Rule::continue_statement as usize)
            && children.len() == 1 && let Node::Label(label) = self.arena.get(children[0])
        {
            let recipe = if rule == Rule::goto_statement as usize { ControlRecipe::Goto(*label) }
                else { ControlRecipe::Continue(Some(*label)) };
            let id = self.arena.unary_with_metadata(children[0], |_| {
                (Node::Control(Control { value: recipe.replay(source), recipe }), Undo::Label)
            }, unwrap_control_label);
            self.arena.push(id);
            return;
        }
        if rule == Rule::statement as usize && let Some(capture) = capture {
            if children.len() == 1 && matches!(self.arena.get(children[0]), Node::Control(_)) {
                let id = self.arena.unary_with_metadata(children[0], |node| {
                    let Node::Control(control) = node else { unreachable!(); };
                    (Node::BodyItem(BodyReceipt { capture, span, control: Some(control) }), Undo::BodyItem)
                }, unwrap_body_item);
                self.arena.push(id);
            } else {
                self.push(Node::BodyItem(BodyReceipt { capture, span, control: None }));
            }
            return;
        }
        if rule == Rule::compound_statement as usize
            && children.iter().all(|id| matches!(self.arena.get(*id), Node::BodyItem(_)))
        {
            let id = self.arena.list(children, |nodes| Node::Body { span,
                items: nodes.into_iter().map(|node| match node {
                    Node::BodyItem(item) => item, _ => unreachable!(),
                }).collect(),
            }, split_body);
            self.arena.push(id);
            return;
        }
        let origin = Origin { anchor: None };
        if rule == Rule::prefix_op as usize && children.is_empty() {
            let op = match &source[span.start..span.end] {
                "+" => Some("+"), "-" => Some("-"), "~" => Some("~"), _ => None,
            };
            if let Some(op) = op { self.push(Node::Prefix(op)); return; }
        }
        if rule == Rule::unary_expression as usize && children.len() == 2
            && let Node::Prefix(op) = self.arena.get(children[0])
            && matches!(self.arena.get(children[1]), Node::Ready { recipe, .. } if recipe.numeric())
        {
            let op = *op;
            let id = self.arena.unary_with_metadata(children[1], |node| {
                let Node::Ready { origin, value, recipe } = node else { unreachable!(); };
                (Node::Ready { origin, value: crate::walker::construct_numeric_prefix(op, value),
                    recipe: Recipe::Prefix { op, operand: Box::new(recipe) } }, Undo::Prefix(origin))
            }, unwrap_prefix);
            self.arena.push(id);
            return;
        }
        if let Some(literal_rule) = Self::literal(rule) {
            if children.is_empty() {
                self.push(Node::Ready {
                    origin,
                    value: crate::walker::construct_literal(literal_rule, &source[span.start..span.end]),
                    recipe: Recipe::Literal { rule: literal_rule, span },
                });
                return;
            }
        } else if rule == Rule::initializer_list as usize
            && children.iter().all(|id| matches!(self.arena.get(*id), Node::Ready { .. }))
        {
            #[cfg(test)]
            if self.reference_construction() {
                self.finish_reference_array(children, span);
                return;
            }
            let id = self.arena.list_with_metadata(children, |children| {
                let mut values = Vec::with_capacity(children.len());
                let mut recipes = Vec::with_capacity(children.len());
                let mut anchors = Vec::new();
                for (index, child) in children.into_iter().enumerate() {
                    let Node::Ready { origin, value, recipe } = child else { unreachable!(); };
                    if let Some(anchor) = origin.anchor { anchors.push((index, anchor)); }
                    values.push(ArrayElement { key: None, value, spread: false, by_ref: false });
                    recipes.push(recipe);
                }
                (Node::Ready { origin: Origin { anchor: Some(span), ..origin },
                    value: Expression::new(ExprKind::Array(values)), recipe: Recipe::Array(recipes) }, Undo::Array(anchors))
            }, split_array);
            self.arena.push(id);
            return;
        } else if Self::wrapper(rule) && children.len() == 1
            && matches!(self.arena.get(children[0]), Node::Ready { .. })
        {
            // A syntax-only wrapper changes neither the AST nor its ownership
            // anchor. Reuse the arena ID; stack edits already journal rollback.
            self.arena.push(children[0]);
            return;
        }
        // Unconsumed ready children remain arena-owned; supported nested brace
        // islands can still be handed off within an unsupported outer region.
        self.push(Node::Unsupported);
    }
}

impl Builder for Direct<'_> {
    fn supports_choice_heads(&self) -> bool { super::choice_heads_enabled() }
    fn observes_rule(&self, rule: usize) -> bool {
        // Identity wrappers do not need frames: their operator/designator/
        // suffix rules are independently observed rejection guards. Conditional
        // and comma expressions retain frames because their punctuation is silent.
        let expression = rule == Rule::initializer_list as usize
            || self.depth > 0 && rule != Rule::initializer as usize
                && rule != Rule::primary_expression as usize && rule != Rule::literal as usize
                && rule != Rule::designated_init as usize && rule != Rule::assignment_expression as usize
                && rule != Rule::binary_expression as usize && rule != Rule::postfix_expression as usize;
        let parent = self.current_frame().map(|(rule, _)| rule);
        !self.literal_active && (expression || Self::body_rule(rule)
            || self.depth == 0 && Self::syntax_rule(rule)
            || self.depth == 0 && [Rule::init_declarator, Rule::declarator,
                Rule::direct_declarator, Rule::initializer, Rule::pointer,
                Rule::array_suffix, Rule::param_suffix, Rule::parameter_decl].iter().any(|candidate| rule == *candidate as usize)
            || self.depth == 0 && rule == Rule::declaration_specifiers as usize
                && matches!(parent, Some(parent) if parent == Rule::parameter_decl as usize
                    || parent == Rule::function_definition as usize)
            || self.depth == 0 && rule == Rule::normal_declaration as usize
                && parent == Some(Rule::function_definition as usize)
            || self.depth == 0 && parent == Some(Rule::declaration_specifiers as usize)
                && [Rule::gcc_attribute, Rule::alignment_specifier, Rule::typeof_specifier,
                    Rule::struct_or_union_spec, Rule::enum_spec].iter().any(|candidate| rule == *candidate as usize)
            || self.depth == 0 && rule == Rule::primary_expression as usize
            || parent == Some(Rule::primary_expression as usize)
                && (Self::literal(rule).is_some() || rule == Rule::ident_name as usize
                    || rule == Rule::expression as usize)
            || rule == Rule::ident_name as usize && matches!(parent,
                Some(parent) if parent == Rule::goto_statement as usize || parent == Rule::continue_statement as usize
                    || parent == Rule::direct_declarator as usize)
            || rule == Rule::ident_name as usize && matches!(parent,
                Some(parent) if parent == Rule::member_suffix as usize || parent == Rule::arrow_suffix as usize)
            || rule == Rule::expression as usize && parent == Some(Rule::return_statement as usize))
    }
    fn observes_capture(&self, rule: usize) -> bool {
        // An admitted initializer cannot contain statements. Do not make its
        // certified complete-island elision depend on unrelated body handles.
        self.depth == 0 && (rule == Rule::statement as usize || rule == Rule::unary_expression as usize
            || rule == Rule::declaration_specifiers as usize || rule == Rule::expression as usize
            || rule == Rule::assignment_expression as usize
            || rule == Rule::ident_name as usize
            || [Rule::declarator, Rule::normal_declaration, Rule::compound_statement]
                .iter().any(|candidate| rule == *candidate as usize)
            || [Rule::for_init, Rule::for_cond, Rule::for_update].iter().any(|candidate| rule == *candidate as usize))
    }
    fn supports_unobserved_failure_prefixes(&self) -> bool { true }
    fn opaque_rule(&self, rule: usize) -> bool {
        self.depth == 0 || Self::literal(rule).is_some()
    }
    fn checkpoint(&self) -> usize { self.arena.checkpoint() + self.state_undo.len() }
    fn rollback(&mut self, mark: usize) {
        while self.state_undo.last().is_some_and(|(arena_mark, _)|
            arena_mark + self.state_undo.len() - 1 >= mark)
        {
            match self.state_undo.pop().unwrap().1 {
                StateUndo::State { depth, literal_active } => {
                    self.depth = depth;
                    self.literal_active = literal_active;
                }
                StateUndo::PopFrame => { self.frames.pop().expect("direct frame push inverse"); }
                StateUndo::PushFrame(frame) => self.frames.push(frame),
            }
        }
        self.arena.rollback(mark - self.state_undo.len());
    }
    fn begin(&mut self, rule: usize, start: usize) {
        // Do not observe/journal an unsupported table's expression descendants.
        // Its nested brace islands remain independently eligible at depth zero.
        if rule == Rule::initializer_list as usize && self.depth == 0 {
            let Some(end) = initializer_end(self.source, start) else { return; };
            // Source-only memo, not transactional state: the ownership query
            // immediately after begin need not scan this same table again.
            self.admitted_end = Some((start, end));
        }
        if rule == Rule::initializer_list as usize { self.set_state(self.depth + 1, false); }
        #[cfg(test)]
        if self.reference_construction() {
            self.push(Node::Frame { rule, start });
        } else {
            self.state_undo.push((self.arena.checkpoint(), StateUndo::PopFrame));
            self.frames.push(Frame { rule, start, base: self.arena.values().len() });
        }
        #[cfg(not(test))]
        {
            self.state_undo.push((self.arena.checkpoint(), StateUndo::PopFrame));
            self.frames.push(Frame { rule, start, base: self.arena.values().len() });
        }
        if Self::literal(rule).is_some() { self.set_state(self.depth, true); }
    }
    fn finish(&mut self, rule: usize, span: Span, source: &str) {
        self.finish_with_capture(rule, span, source, None);
    }
    fn try_finish_capture(&mut self, rule: usize, span: Span, source: &str,
        capture: Option<CaptureHandle>) -> Result<(), grameo::builder::BuildError>
    {
        self.finish_with_capture(rule, span, source, capture);
        Ok(())
    }
    fn reduce(&mut self, _: usize, fixity: grameo::pratt::Fixity, span: Span) {
        if self.depth != 0 || fixity != grameo::pratt::Fixity::InfixLeft
            || self.current_frame().map(|(rule, _)| rule) != Some(Rule::binary_expression as usize)
        { return; }
        let right = self.arena.pop().expect("Pratt right operand");
        let left = self.arena.pop().expect("Pratt left operand");
        let supported = |node: &Node| matches!(node,
            Node::Syntax { recipe, .. } if matches!(&**recipe, SyntaxRecipe::Unary { .. }))
            || matches!(node, Node::Binary { span: None, .. });
        if !supported(self.arena.get(left)) || !supported(self.arena.get(right)) {
            self.push(Node::Unsupported);
            return;
        }
        let op = super::owned::operator_kind(&self.source[span.start..span.end]);
        let id = self.arena.list_with_metadata(vec![left, right], |nodes| {
            let mut nodes = nodes.into_iter();
            let (left, lrecipe, lorigin) = binary_child(nodes.next().unwrap());
            let (right, rrecipe, rorigin) = binary_child(nodes.next().unwrap());
            (Node::Binary { span: None,
                value: Expression::new(ExprKind::Binary { op, left: Box::new(left), right: Box::new(right) }),
                recipe: Rc::new(BinaryRecipe::Infix { op, left: lrecipe, right: rrecipe }) },
                Undo::Binary([lorigin, rorigin]))
        }, split_binary);
        self.arena.push(id);
    }
    fn committed(&mut self, _: &str, _: usize) {
        self.committed = true;
        self.arena.commit();
        self.state_undo.clear();
    }
    fn supports_owned_capture_boundary(&self, rule: usize) -> bool {
        self.elide && self.depth > 0 && rule == Rule::initializer_list as usize
    }
    fn owned_capture_end(&self, rule: usize, start: usize, source: &str) -> Option<usize> {
        if rule != Rule::initializer_list as usize || source.as_ptr() != self.source.as_ptr()
            || source.len() != self.source.len() { return None; }
        if !self.reference_construction()
            && let Some((from, end)) = self.admitted_end && from == start { return Some(end); }
        initializer_end(source, start)
    }
}

/// Conservative admission/end prediction only; the original grammar still
/// recognizes every token and owns diagnostics, budgets and exact-end fallback.
fn initializer_end(source: &str, start: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    if bytes.get(start) != Some(&b'{') { return None; }
    let mut stack = vec![(b'}', true)];
    let mut at = start + 1;
    let mut string_tail = false;
    let mut numeric_prefix = false;
    while let Some(&byte) = bytes.get(at) {
        if byte.is_ascii_whitespace() { at += 1; continue; }
        if source[at..].starts_with("/*") {
            at += 2 + source[at + 2..].find("*/")? + 2;
            continue;
        }
        if source[at..].starts_with("//") {
            at += source[at..].find('\n')?;
            continue;
        }
        match byte {
            b'{' | b'(' => {
                if !stack.last()?.1 { return None; }
                if byte == b'{' && numeric_prefix { return None; }
                stack.push((if byte == b'{' { b'}' } else { b')' }, true));
                at += 1;
                string_tail = false;
            }
            b'}' | b')' => {
                let (end, needs_value) = stack.pop()?;
                if end != byte || byte == b')' && needs_value { return None; }
                at += 1;
                if stack.is_empty() { return Some(at); }
                stack.last_mut()?.1 = false;
                string_tail = false;
                numeric_prefix = false;
            }
            b',' => {
                let frame = stack.last_mut()?;
                if frame.0 != b'}' || frame.1 { return None; }
                frame.1 = true;
                at += 1;
                string_tail = false;
                numeric_prefix = false;
            }
            b'+' | b'-' | b'~' => {
                if !stack.last()?.1 || matches!(byte, b'+' | b'-')
                    && bytes.get(at + 1) == Some(&byte) { return None; }
                numeric_prefix = true;
                at += 1;
            }
            b'\'' | b'"' => {
                if byte == b'"' && numeric_prefix { return None; }
                if !stack.last()?.1 && !(byte == b'"' && string_tail) { return None; }
                let quote = byte;
                at += 1;
                loop {
                    match *bytes.get(at)? {
                        b'\\' => { at += 1; at += source[at..].chars().next()?.len_utf8(); }
                        end if end == quote => { at += 1; break; }
                        _ => { at += source[at..].chars().next()?.len_utf8(); }
                    }
                }
                stack.last_mut()?.1 = false;
                string_tail = quote == b'"';
                numeric_prefix = false;
            }
            b'0'..=b'9' | b'.' => {
                if !stack.last()?.1 { return None; }
                let from = at;
                // e/E can be a hex integer digit, not an exponent sign anchor.
                // Preserve even legacy PEG acceptance when it differs from C's
                // preprocessing-number maximal munch.
                let hexadecimal = source[from..].starts_with("0x")
                    || source[from..].starts_with("0X");
                while let Some(&ch) = bytes.get(at) {
                    if ch.is_ascii_alphanumeric() || ch == b'.'
                        || matches!(ch, b'+' | b'-') && at > from
                            && if hexadecimal { matches!(bytes[at - 1], b'p' | b'P') }
                                else { matches!(bytes[at - 1], b'e' | b'E') }
                    { at += 1; } else { break; }
                }
                if at == from { return None; }
                stack.last_mut()?.1 = false;
                string_tail = false;
                numeric_prefix = false;
            }
            _ => {
                if let Some(prefix) = ["u8", "L", "u", "U"].into_iter().find(|prefix|
                    source[at..].strip_prefix(prefix).is_some_and(|tail|
                        tail.starts_with(['\'', '"'])))
                {
                    if !stack.last()?.1 && !string_tail { return None; }
                    at += prefix.len(); continue;
                }
                if !stack.last()?.1 { return None; }
                let word = ["true", "false"].into_iter().find(|word|
                    source[at..].strip_prefix(word).is_some_and(|tail|
                        tail.as_bytes().first().is_none_or(|ch|
                            !ch.is_ascii_alphanumeric() && *ch != b'_')))?;
                at += word.len();
                stack.last_mut()?.1 = false;
                string_tail = false;
                numeric_prefix = false;
            }
        }
    }
    None
}

#[cfg(test)]
#[path = "direct_tests.rs"]
mod tests;

#[cfg(test)]
#[path = "direct_param_tests.rs"]
mod param_tests;

#[cfg(test)]
#[path = "direct_if_tests.rs"]
mod if_tests;

#[cfg(test)]
#[path = "direct_loop_tests.rs"]
mod loop_tests;

#[cfg(test)]
#[path = "direct_for_tests.rs"]
mod for_tests;

#[cfg(test)]
#[path = "direct_switch_tests.rs"]
mod switch_tests;

#[cfg(test)]
#[path = "direct_call_tests.rs"]
mod call_tests;

#[cfg(test)]
#[path = "direct_place_tests.rs"]
mod place_tests;

#[cfg(test)]
#[path = "direct_function_tests.rs"]
mod function_tests;

#[cfg(test)]
#[path = "direct_return_tests.rs"]
mod return_tests;
