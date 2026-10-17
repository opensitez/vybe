//! C-private whole-function goto preparation, enabled only by an internal opt-in.
//! Call only before any block/case has lowered its labels. Leading declarations
//! remain unchanged. The caller certifies only later plain scalar bindings whose
//! reservation has no observable initialization.
use std::collections::{HashMap, HashSet};
use vybe_ast::{BinOp, BindingPattern, BitLane, BreakTarget, CaseCondition, ContinueTarget,
    ExprKind, Expression, IntOp, Literal, Span, Statement, StmtKind, StorageAccess,
    SwitchCase, VarDeclKind, VarDeclarator, LambdaBody};

#[derive(Debug)]
pub(crate) struct LowerError {
    pub(crate) reason: String,
    pub(crate) span: Span,
}

type Id = usize;

enum Edge {
    Next(Id),
    Branch(Expression, Id, Id),
    Select(Expression, Vec<(Vec<CaseCondition>, Id)>, Id),
    Goto(String),
    Stop,
    Return,
}

struct Node { statements: Vec<Statement>, edge: Edge, span: Span }

#[derive(Clone, Copy, Default)]
struct Targets { break_to: Option<Id>, continue_to: Option<Id> }

struct Graph<'a> {
    nodes: Vec<Node>,
    labels: HashMap<String, Id>,
    bindings: HashSet<String>,
    prelude: Vec<Statement>,
    split_safe: &'a HashSet<String>,
}

fn expression(kind: ExprKind) -> Expression { Expression::new(kind) }
fn ident(name: &str) -> Expression { expression(ExprKind::Ident(name.into())) }
fn integer(value: i64) -> Expression { expression(ExprKind::Lit(Literal::Int(value))) }

fn assign(name: &str, value: Expression, span: Span) -> Statement {
    Statement::with_span(StmtKind::Expr(expression(ExprKind::Assign {
        target: Box::new(ident(name)), value: Box::new(value),
    })), span)
}

fn error(span: &Span, reason: impl Into<String>) -> LowerError {
    LowerError { span: span.clone(), reason: reason.into() }
}

/// Transactional preparation: errors retain the caller's entire input AST.
/// `split_safe` must certify uniquely renamed, nonescaping ordinary scalar
/// locals: reserving a zero/default slot must neither allocate nor run effects.
/// The maximal leading Let/Var declaration prefix is never split or reserved:
/// its original storage and initializers remain at function entry exactly once.
/// Unknown/volatile/atomic/aggregate/VLA/static/const bindings are not inferred
/// safe from an arbitrary initializer or from the identifier spelling.
pub(crate) fn lower_function_gotos(
    body: &[Statement], split_safe: &HashSet<String>, reserved_names: &HashSet<String>,
) -> Result<Vec<Statement>, LowerError> {
    validate_callable_control(body)?;
    lower_validated_gotos(body, split_safe, reserved_names)
}

pub(crate) fn validate_callable_control(body: &[Statement]) -> Result<(), LowerError> {
    let mut embedded = None;
    for statement in body {
        if matches!(&statement.kind, StmtKind::FunctionDecl { .. }) { continue; }
        statement.walk_exprs(&mut |value| {
            let unprepared = match &value.kind {
                ExprKind::Lambda { body: LambdaBody::Block(body), .. } => contains_unprepared_control(body, true),
                ExprKind::FunctionExpr(function) => match &function.kind {
                    StmtKind::FunctionDecl { body, .. } => contains_unprepared_control(body, true), _ => false,
                },
                _ => false,
            };
            if unprepared {
                embedded = Some(error(&value.span, "control label/goto in expression-owned callable requires separate scope preparation"));
            }
        });
    }
    if let Some(error) = embedded { return Err(error); }
    Ok(())
}

fn lower_validated_gotos(
    body: &[Statement], split_safe: &HashSet<String>, reserved_names: &HashSet<String>,
) -> Result<Vec<Statement>, LowerError> {
    if !contains_goto(body) { return Ok(body.to_vec()); }
    // No edge can enter before the first label. Preserve leading declarations
    // exactly, including initialization/storage, outside the dispatcher.
    let prefix_len = body.iter().take_while(|statement| matches!(&statement.kind,
        StmtKind::VarDecl { kind: VarDeclKind::Let | VarDeclKind::Var, .. })).count();
    let (prefix, remaining) = body.split_at(prefix_len);
    let mut prefix_bindings = HashSet::new();
    for statement in prefix {
        let StmtKind::VarDecl { declarations, .. } = &statement.kind else { unreachable!(); };
        for declaration in declarations {
            let BindingPattern::Ident(name) = &declaration.pattern else {
                return Err(error(&statement.span, "destructured binding in leading C declaration prefix"));
            };
            if !prefix_bindings.insert(name.clone()) {
                return Err(error(&statement.span, format!("duplicate leading C binding: {name}")));
            }
        }
    }
    let mut graph = Graph { nodes: Vec::new(), labels: HashMap::new(),
        bindings: prefix_bindings, prelude: prefix.to_vec(), split_safe };
    let end = graph.node(Vec::new(), Edge::Stop, Span::default());
    let entry = graph.sequence(remaining, end, Targets::default())?;
    for node in &mut graph.nodes {
        if let Edge::Goto(label) = &node.edge {
            let target = graph.labels.get(label).ok_or_else(||
                error(&node.span, format!("unresolved function-scope C label: {label}")))?;
            node.edge = Edge::Next(*target);
        }
    }
    let mut used = reserved_names.clone();
    used.extend(graph.bindings.iter().cloned());
    used.extend(graph.labels.keys().cloned());
    for statement in body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Ident(name) = &value.kind { used.insert(name.clone()); }
        });
    }
    let fresh = |base: &str, used: &mut HashSet<String>| {
        let mut index = 0;
        loop {
            let candidate = format!("{base}_{index}");
            if used.insert(candidate.clone()) { break candidate; }
            index += 1;
        }
    };
    let pc = fresh("__c_function_goto_pc", &mut used);
    let dispatch = fresh("__c_function_goto_dispatch", &mut used);
    let jump = |target: Id, span: &Span| vec![
        assign(&pc, integer(target as i64), span.clone()),
        Statement::with_span(StmtKind::Continue(ContinueTarget::Label(dispatch.clone())), span.clone()),
    ];
    let mut cases = Vec::with_capacity(graph.nodes.len());
    for (id, node) in graph.nodes.into_iter().enumerate() {
        let mut statements = node.statements;
        match node.edge {
            Edge::Next(next) => statements.extend(jump(next, &node.span)),
            Edge::Branch(condition, yes, no) => {
                statements.push(Statement::with_span(StmtKind::If {
                    cond: condition, then_body: jump(yes, &node.span), elifs: Vec::new(),
                    else_body: Some(jump(no, &node.span)),
                }, node.span));
            }
            Edge::Select(selector, choices, default) => {
                statements.push(Statement::with_span(StmtKind::Switch {
                    expr: selector,
                    cases: choices.into_iter().map(|(conditions, entry)| SwitchCase {
                        conditions, body: jump(entry, &node.span),
                    }).collect(),
                    default: Some(jump(default, &node.span)),
                }, node.span));
            }
            Edge::Stop => {
                statements.push(assign(&pc, integer(-1), node.span.clone()));
                statements.push(Statement::with_span(
                    StmtKind::Continue(ContinueTarget::Label(dispatch.clone())), node.span));
            }
            Edge::Return => {}
            Edge::Goto(_) => unreachable!("all label references resolved before publication"),
        }
        cases.push(SwitchCase { conditions: vec![CaseCondition::Value(integer(id as i64))], body: statements });
    }
    graph.prelude.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator { pattern: BindingPattern::Ident(pc.clone()),
            type_hint: Some("int".into()), init: Some(integer(entry as i64)), array_bounds: None, with_events: false }],
        kind: VarDeclKind::Let,
    }));
    graph.prelude.push(Statement::new(StmtKind::Labeled {
        label: dispatch,
        body: Box::new(Statement::new(StmtKind::While {
            cond: expression(ExprKind::Binary {
                op: BinOp::Integer(IntOp::GeS, BitLane::W32),
                left: Box::new(ident(&pc)), right: Box::new(integer(0)),
            }),
            body: vec![Statement::new(StmtKind::Switch {
                expr: ident(&pc), cases,
                default: Some(vec![assign(&pc, integer(-1), Span::default())]),
            })],
            else_body: None,
        })),
    }));
    Ok(graph.prelude)
}

impl Graph<'_> {
    fn node(&mut self, statements: Vec<Statement>, edge: Edge, span: Span) -> Id {
        let id = self.nodes.len();
        self.nodes.push(Node { statements, edge, span });
        id
    }

    fn sequence(&mut self, body: &[Statement], mut next: Id, targets: Targets) -> Result<Id, LowerError> {
        for statement in body.iter().rev() { next = self.statement(statement, next, targets)?; }
        Ok(next)
    }

    fn statement(&mut self, statement: &Statement, next: Id, targets: Targets) -> Result<Id, LowerError> {
        let span = &statement.span;
        match &statement.kind {
            StmtKind::Label(label) => {
                if self.labels.insert(label.clone(), next).is_some() {
                    return Err(error(span, format!("duplicate function-scope C label: {label}")));
                }
                Ok(next)
            }
            StmtKind::GoTo(label) => Ok(self.node(Vec::new(), Edge::Goto(label.clone()), span.clone())),
            StmtKind::Block(body) => self.sequence(body, next, targets),
            StmtKind::If { cond, then_body, elifs, else_body } => {
                let mut otherwise = self.sequence(else_body.as_deref().unwrap_or(&[]), next, targets)?;
                for (condition, body) in elifs.iter().rev() {
                    let yes = self.sequence(body, next, targets)?;
                    otherwise = self.node(Vec::new(), Edge::Branch(condition.clone(), yes, otherwise), span.clone());
                }
                let yes = self.sequence(then_body, next, targets)?;
                Ok(self.node(Vec::new(), Edge::Branch(cond.clone(), yes, otherwise), span.clone()))
            }
            StmtKind::While { cond, body, else_body: None } => {
                let test = self.node(Vec::new(), Edge::Stop, span.clone());
                let body = self.sequence(body, test, Targets { break_to: Some(next), continue_to: Some(test) })?;
                self.nodes[test].edge = Edge::Branch(cond.clone(), body, next);
                Ok(test)
            }
            StmtKind::DoWhile { body, cond, until } => {
                let test = self.node(Vec::new(), Edge::Stop, span.clone());
                let body = self.sequence(body, test, Targets { break_to: Some(next), continue_to: Some(test) })?;
                self.nodes[test].edge = if *until { Edge::Branch(cond.clone(), next, body) }
                    else { Edge::Branch(cond.clone(), body, next) };
                Ok(body)
            }
            StmtKind::For { init, cond, update, body } => {
                let test = self.node(Vec::new(), Edge::Stop, span.clone());
                let update = match update {
                    Some(value) => self.node(vec![Statement::with_span(StmtKind::Expr(value.clone()), span.clone())],
                        Edge::Next(test), span.clone()),
                    None => test,
                };
                let body = self.sequence(body, update, Targets { break_to: Some(next), continue_to: Some(update) })?;
                self.nodes[test].edge = match cond {
                    Some(value) => Edge::Branch(value.clone(), body, next), None => Edge::Next(body),
                };
                match init { Some(value) => self.statement(value, test, targets), None => Ok(test) }
            }
            StmtKind::Switch { expr, cases, default } => {
                let switch_targets = Targets { break_to: Some(next), ..targets };
                let mut default_entry = match default {
                    Some(body) => self.sequence(body, next, switch_targets)?, None => next,
                };
                let mut following = default_entry;
                let mut choices = Vec::new();
                let mut has_default = default.is_some();
                for case in cases.iter().rev() {
                    let entry = self.sequence(&case.body, following, switch_targets)?;
                    following = entry;
                    if case.conditions.is_empty() {
                        if has_default { return Err(error(span, "multiple switch defaults")); }
                        has_default = true;
                        default_entry = entry;
                    } else { choices.push((case.conditions.clone(), entry)); }
                }
                choices.reverse();
                Ok(self.node(Vec::new(), Edge::Select(expr.clone(), choices, default_entry), span.clone()))
            }
            StmtKind::Break(BreakTarget::Implicit) => targets.break_to
                .map(|target| self.node(Vec::new(), Edge::Next(target), span.clone()))
                .ok_or_else(|| error(span, "break without a containing C loop/switch")),
            StmtKind::Continue(ContinueTarget::Implicit) => targets.continue_to
                .map(|target| self.node(Vec::new(), Edge::Next(target), span.clone()))
                .ok_or_else(|| error(span, "continue without a containing C loop")),
            StmtKind::VarDecl { declarations, kind } => {
                if !matches!(kind, VarDeclKind::Let | VarDeclKind::Var) {
                    return Err(error(span, "uncertified declaration storage/lifetime in C goto region"));
                }
                let mut initializers = Vec::new();
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        return Err(error(span, "destructured binding in C goto region"));
                    };
                    if !self.split_safe.contains(name) || !self.bindings.insert(name.clone())
                        || declaration.array_bounds.is_some() || declaration.with_events
                        || declaration.type_hint.as_ref().is_some_and(|hint|
                            hint.spelling().contains('[')
                                || matches!(hint.access, Some(StorageAccess::Volatile | StorageAccess::Atomic | StorageAccess::Unknown)))
                    {
                        return Err(error(span, format!("binding lacks unique ordinary scalar reservation proof: {name}")));
                    }
                    let mut reserved = declaration.clone();
                    reserved.init = None;
                    self.prelude.push(Statement::with_span(StmtKind::VarDecl {
                        declarations: vec![reserved], kind: kind.clone(),
                    }, span.clone()));
                    if let Some(value) = &declaration.init {
                        let value = match &declaration.type_hint {
                            Some(hint) if hint.converts() => expression(ExprKind::Cast {
                                type_name: hint.spelling().into(), expr: Box::new(value.clone()),
                            }),
                            _ => value.clone(),
                        };
                        initializers.push(assign(name, value, span.clone()));
                    }
                }
                Ok(self.node(initializers, Edge::Next(next), span.clone()))
            }
            StmtKind::Return(_) => Ok(self.node(vec![statement.clone()], Edge::Return, span.clone())),
            StmtKind::Expr(_) | StmtKind::Assign { .. } | StmtKind::CompoundAssign { .. } | StmtKind::Empty =>
                Ok(self.node(vec![statement.clone()], Edge::Next(next), span.clone())),
            _ => Err(error(span, "unsupported structured scope/control in C goto region")),
        }
    }
}

fn contains_goto(body: &[Statement]) -> bool {
    contains_unprepared_control(body, false)
}

pub(crate) fn contains_unprepared_control(body: &[Statement], labels: bool) -> bool {
    body.iter().any(|statement| match &statement.kind {
        StmtKind::GoTo(_) => true,
        StmtKind::Label(_) if labels => true,
        StmtKind::Block(body) | StmtKind::DoWhile { body, .. } | StmtKind::With { body, .. }
            | StmtKind::Using { body, .. } | StmtKind::Lock { body, .. }
            | StmtKind::NamespaceDecl { body, .. } => contains_unprepared_control(body, labels),
        StmtKind::If { then_body, elifs, else_body, .. } => contains_unprepared_control(then_body, labels)
            || elifs.iter().any(|(_, body)| contains_unprepared_control(body, labels))
            || else_body.as_ref().is_some_and(|body| contains_unprepared_control(body, labels)),
        StmtKind::Switch { cases, default, .. } => cases.iter().any(|case| contains_unprepared_control(&case.body, labels))
            || default.as_ref().is_some_and(|body| contains_unprepared_control(body, labels)),
        StmtKind::Labeled { body, .. } => labels || contains_unprepared_control(std::slice::from_ref(body.as_ref()), labels),
        StmtKind::Try { body, catches, else_body, finally } => contains_unprepared_control(body, labels)
            || catches.iter().any(|catch| contains_unprepared_control(&catch.body, labels))
            || [else_body, finally].into_iter().any(|body| body.as_ref().is_some_and(|body| contains_unprepared_control(body, labels))),
        StmtKind::ForIn { body, else_body, .. } => contains_unprepared_control(body, labels)
            || else_body.as_ref().is_some_and(|body| contains_unprepared_control(body, labels)),
        StmtKind::While { body, else_body, .. } => contains_unprepared_control(body, labels)
            || else_body.as_ref().is_some_and(|body| contains_unprepared_control(body, labels)),
        StmtKind::For { init, body, .. } => contains_unprepared_control(body, labels)
            || init.as_ref().is_some_and(|init| contains_unprepared_control(std::slice::from_ref(init.as_ref()), labels)),
        // Nested functions belong to their own function-scope preparation.
        StmtKind::FunctionDecl { .. } => false,
        _ => false,
    })
}

#[cfg(test)]
#[path = "goto_lowering_tests.rs"]
pub(crate) mod tests;
