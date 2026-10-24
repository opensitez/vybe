use super::{BindingProof, WordInterpretation};
use vybe_ast::{
    ArrayPatternElem, BinOp, BindingPattern, BitLane, Directives, ExprKind, Expression, Literal,
    PlaceExpr, Statement, StmtKind, StorageAccess, TypeBinding, TypeHint, UnaryOp, VarDeclKind,
    WordLocalTransport,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Rejection {
    UndeclaredOrUninitialized,
    IncompatibleProducer,
    AddressOrReferenceEscape,
    Capture,
    ShadowedBinding,
    UnsupportedExpression,
    UnsupportedControlFlow,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RejectedBinding {
    pub name: String,
    pub reason: Rejection,
}

#[derive(Debug, Default)]
pub struct SubsetProof {
    pub eligible: Vec<BindingProof>,
    pub rejected: Vec<RejectedBinding>,
}

/// Select bindings, not whole bodies. Unrelated pointers/memory retain their
/// original ABI. Cast resolution must certify an actual 32-bit integer cast;
/// a spelling, inferred Number result or arbitrary helper call is not proof.
pub fn prove_subset(
    directives: &Directives,
    body: &[Statement],
    resolve: impl Fn(&TypeHint) -> Option<WordInterpretation>,
    resolve_cast: impl Fn(&str) -> Option<WordInterpretation>,
) -> SubsetProof {
    let mut result = SubsetProof::default();
    if directives.word_local_transport() != WordLocalTransport::ProvenW32 {
        return result;
    }
    for statement in body {
        let StmtKind::VarDecl { declarations, kind } = &statement.kind else {
            continue;
        };
        if matches!(kind, VarDeclKind::Static | VarDeclKind::FunctionScoped) {
            continue;
        }
        for declaration in declarations {
            let BindingPattern::Ident(name) = &declaration.pattern else {
                continue;
            };
            let Some(hint) = &declaration.type_hint else {
                continue;
            };
            if hint.binding != TypeBinding::Checked
                || hint.access != Some(StorageAccess::Ordinary)
                || declaration.with_events
                || declaration.array_bounds.is_some()
            {
                continue;
            }
            let Some(interpretation) = resolve(hint) else {
                continue;
            };
            let mut analysis = Analysis {
                name,
                interpretation,
                resolve_cast: &resolve_cast,
                declared: false,
            };
            let mut initialized = false;
            match analysis.body(body, &mut initialized, true) {
                Ok(()) => result.eligible.push(BindingProof {
                    name: name.clone(),
                    interpretation,
                }),
                Err(reason) => result.rejected.push(RejectedBinding {
                    name: name.clone(),
                    reason,
                }),
            }
        }
    }
    result
}

#[derive(Clone, Copy)]
enum Producer {
    Unknown,
    Literal(i64),
    Signed,
    Unsigned,
}

struct Analysis<'a, F> {
    name: &'a str,
    interpretation: WordInterpretation,
    resolve_cast: &'a F,
    declared: bool,
}

fn contains(value: &Expression, name: &str) -> bool {
    let mut found = false;
    value.walk_exprs(&mut |value| match &value.kind {
        ExprKind::Ident(other) => found |= other == name,
        ExprKind::RefOf(place) => {
            found |= matches!(&**place, PlaceExpr::Ident(other) if other == name)
        }
        ExprKind::Lambda { captures, .. } => found |= captures.iter().any(|other| other == name),
        _ => {}
    });
    found
}

fn pattern_contains(pattern: &BindingPattern, name: &str) -> bool {
    match pattern {
        BindingPattern::Ident(other) => other == name,
        BindingPattern::Object(properties) => properties.iter().any(|property| {
            property
                .value
                .as_ref()
                .map_or(property.key == name, |value| pattern_contains(value, name))
        }),
        BindingPattern::Array(elements) => elements.iter().any(|element| match element {
            ArrayPatternElem::Pattern(pattern, _) => pattern_contains(pattern, name),
            ArrayPatternElem::Rest(other) => other == name,
            ArrayPatternElem::Hole => false,
        }),
    }
}

impl<F: Fn(&str) -> Option<WordInterpretation>> Analysis<'_, F> {
    fn representation(&self) -> Producer {
        match self.interpretation {
            WordInterpretation::Signed => Producer::Signed,
            WordInterpretation::Unsigned => Producer::Unsigned,
        }
    }

    fn compatible(&self, producer: Producer) -> bool {
        match (self.interpretation, producer) {
            (WordInterpretation::Signed, Producer::Signed)
            | (WordInterpretation::Unsigned, Producer::Unsigned) => true,
            (WordInterpretation::Signed, Producer::Literal(n)) => i32::try_from(n).is_ok(),
            (WordInterpretation::Unsigned, Producer::Literal(n)) => u32::try_from(n).is_ok(),
            _ => false,
        }
    }

    fn producer(&self, value: &Expression) -> Producer {
        match &value.kind {
            ExprKind::Lit(Literal::Int(n)) => Producer::Literal(*n),
            ExprKind::Ident(name) if name == self.name => self.representation(),
            ExprKind::Binary {
                op: BinOp::Integer(op, BitLane::W32),
                ..
            } if !op.is_comparison() => Producer::Signed,
            ExprKind::Cast { expr, type_name } => {
                let finite_word = match self.producer(expr) {
                    Producer::Signed | Producer::Unsigned => true,
                    Producer::Literal(n) => {
                        (i64::from(i32::MIN)..=i64::from(u32::MAX)).contains(&n)
                    }
                    Producer::Unknown => false,
                };
                if !finite_word {
                    return Producer::Unknown;
                }
                match (self.resolve_cast)(type_name) {
                    Some(WordInterpretation::Signed) => Producer::Signed,
                    Some(WordInterpretation::Unsigned) => Producer::Unsigned,
                    None => Producer::Unknown,
                }
            }
            ExprKind::Sequence(values) => values
                .last()
                .map_or(Producer::Unknown, |value| self.producer(value)),
            _ => Producer::Unknown,
        }
    }

    fn expression(&mut self, value: &Expression, initialized: &mut bool) -> Result<(), Rejection> {
        if !contains(value, self.name) {
            return Ok(());
        }
        match &value.kind {
            ExprKind::Ident(_) => {
                if !self.declared || !*initialized {
                    return Err(Rejection::UndeclaredOrUninitialized);
                }
            }
            ExprKind::Assign { target, value } => {
                if matches!(&target.kind, ExprKind::Ident(name) if name == self.name) {
                    if !self.declared {
                        return Err(Rejection::UndeclaredOrUninitialized);
                    }
                    self.expression(value, initialized)?;
                    if !self.compatible(self.producer(value)) {
                        return Err(Rejection::IncompatibleProducer);
                    }
                    *initialized = true;
                } else {
                    self.expression(target, initialized)?;
                    self.expression(value, initialized)?;
                }
            }
            ExprKind::Sequence(values) => {
                for value in values {
                    self.expression(value, initialized)?;
                }
            }
            ExprKind::Binary { left, right, op } => {
                self.expression(left, initialized)?;
                let before = *initialized;
                self.expression(right, initialized)?;
                if matches!(op, BinOp::And | BinOp::Or) {
                    *initialized &= before;
                }
            }
            ExprKind::Ternary { cond, then, else_ } => {
                self.expression(cond, initialized)?;
                let mut other = *initialized;
                self.expression(then, initialized)?;
                self.expression(else_, &mut other)?;
                *initialized &= other;
            }
            ExprKind::Member { object, .. } => self.expression(object, initialized)?,
            ExprKind::Index { object, index, .. } => {
                self.expression(object, initialized)?;
                self.expression(index, initialized)?;
            }
            ExprKind::Call {
                callee,
                args,
                optional: false,
            } => {
                self.expression(callee, initialized)?;
                for arg in args {
                    if arg.by_ref && contains(&arg.value, self.name) {
                        return Err(Rejection::AddressOrReferenceEscape);
                    }
                    self.expression(&arg.value, initialized)?;
                }
            }
            ExprKind::Unary {
                op: UnaryOp::AddrOf,
                ..
            }
            | ExprKind::RefOf(_) => return Err(Rejection::AddressOrReferenceEscape),
            ExprKind::Unary {
                op: UnaryOp::PostInc | UnaryOp::PreInc | UnaryOp::PostDec | UnaryOp::PreDec,
                expr,
            } if matches!(&expr.kind, ExprKind::Ident(name) if name == self.name) => {
                return Err(Rejection::IncompatibleProducer);
            }
            ExprKind::Unary { expr, .. }
            | ExprKind::Cast { expr, .. }
            | ExprKind::IsType { expr, .. }
            | ExprKind::TypeOf(expr)
            | ExprKind::RefLoad(expr) => self.expression(expr, initialized)?,
            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) => return Err(Rejection::Capture),
            _ => return Err(Rejection::UnsupportedExpression),
        }
        Ok(())
    }

    fn body(
        &mut self,
        body: &[Statement],
        initialized: &mut bool,
        top: bool,
    ) -> Result<(), Rejection> {
        for statement in body {
            match &statement.kind {
                StmtKind::VarDecl { declarations, .. } => {
                    for declaration in declarations {
                        if pattern_contains(&declaration.pattern, self.name)
                            && !matches!(&declaration.pattern, BindingPattern::Ident(name) if name == self.name)
                        {
                            return Err(Rejection::ShadowedBinding);
                        }
                        if matches!(&declaration.pattern, BindingPattern::Ident(name) if name == self.name)
                        {
                            if !top || self.declared {
                                return Err(Rejection::ShadowedBinding);
                            }
                            self.declared = true;
                            if let Some(value) = &declaration.init {
                                self.expression(value, initialized)?;
                                if !self.compatible(self.producer(value)) {
                                    return Err(Rejection::IncompatibleProducer);
                                }
                                *initialized = true;
                            }
                        } else if let Some(value) = &declaration.init {
                            self.expression(value, initialized)?;
                        }
                    }
                }
                StmtKind::Expr(value) | StmtKind::Return(Some(value)) => {
                    self.expression(value, initialized)?
                }
                StmtKind::Return(None) => {}
                StmtKind::Block(body) => self.body(body, initialized, false)?,
                StmtKind::If {
                    cond,
                    then_body,
                    elifs,
                    else_body,
                } => {
                    self.expression(cond, initialized)?;
                    let before = *initialized;
                    self.body(then_body, initialized, false)?;
                    for (cond, body) in elifs {
                        let mut arm = before;
                        self.expression(cond, &mut arm)?;
                        self.body(body, &mut arm, false)?;
                        *initialized &= arm;
                    }
                    let mut other = before;
                    if let Some(body) = else_body {
                        self.body(body, &mut other, false)?;
                    }
                    *initialized &= other;
                }
                StmtKind::While {
                    cond,
                    body,
                    else_body: None,
                } => {
                    self.expression(cond, initialized)?;
                    let mut iteration = *initialized;
                    self.body(body, &mut iteration, false)?;
                }
                StmtKind::DoWhile {
                    body,
                    cond,
                    until: false,
                } => {
                    let mut iteration = *initialized;
                    self.body(body, &mut iteration, false)?;
                    self.expression(cond, &mut iteration)?;
                    // Conservative: do not export first-iteration facts.
                }
                StmtKind::For {
                    init,
                    cond,
                    update,
                    body,
                } => {
                    if let Some(init) = init {
                        self.body(std::slice::from_ref(init.as_ref()), initialized, false)?;
                    }
                    if let Some(cond) = cond {
                        self.expression(cond, initialized)?;
                    }
                    let mut iteration = *initialized;
                    self.body(body, &mut iteration, false)?;
                    if let Some(update) = update {
                        self.expression(update, &mut iteration)?;
                    }
                }
                // Unknown control flow cannot establish dominance even when
                // the statement itself does not mention this candidate.
                _ => return Err(Rejection::UnsupportedControlFlow),
            }
        }
        Ok(())
    }
}

/// A binding admitted to physical storage in its declared lane.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NativeBindingProof {
    pub name: String,
    pub lane: BitLane,
    pub interpretation: WordInterpretation,
}

#[derive(Debug, Default)]
pub struct NativeSubsetProof {
    pub eligible: Vec<NativeBindingProof>,
    pub rejected: Vec<RejectedBinding>,
}

/// The native-storage lane. A name is eligible when EVERY binding of that name
/// in `params` and `body` (at any depth, so loop-header declarations count) is
/// a checked ordinary declaration resolving to one lane and interpretation,
/// and every access to it is a plain read, a whole-binding assignment or its
/// declaration. Address/reference escape, capture, pattern binding, in-place
/// increments and statements outside the modeled set reject the name.
/// Producers are not certified here: the emitter converts each stored value
/// to the declared lane at its store, which is the declaration's own Checked
/// guarantee, so any producer of the declared type is admissible.
/// W32 analysis ([`prove_subset`]) is unaffected.
pub fn prove_native_subset(
    directives: &Directives,
    params: &[vybe_ast::Param],
    body: &[Statement],
    resolve: impl Fn(&TypeHint) -> Option<(BitLane, WordInterpretation)>,
) -> NativeSubsetProof {
    let mut result = NativeSubsetProof::default();
    if directives.word_local_transport() != WordLocalTransport::ProvenNative {
        return result;
    }
    let mut native = Native {
        resolve: &resolve,
        candidates: Vec::new(),
        rejected: Vec::new(),
    };
    for param in params {
        let declared = (param.pass_by == vybe_ast::PassBy::Value
            && !param.is_rest
            && !param.is_kwargs
            && param.default.is_none())
        .then_some(param.type_hint.as_ref())
        .flatten();
        native.declare(&param.name, declared);
    }
    native.body(body);
    for (name, lane, interpretation) in native.candidates {
        match native.rejected.iter().find(|(other, _)| *other == name) {
            Some((_, reason)) => result.rejected.push(RejectedBinding { name, reason: *reason }),
            None => result.eligible.push(NativeBindingProof { name, lane, interpretation }),
        }
    }
    result
}

/// An assignment target that binds names it mentions (a tuple or array
/// pattern). An element or field target only READS the names in its object
/// and index expressions.
fn writes_through_pattern(target: &Expression) -> bool {
    !matches!(
        &target.kind,
        ExprKind::Ident(_) | ExprKind::Index { .. } | ExprKind::Member { .. } | ExprKind::RefLoad(_)
    ) && !matches!(&target.kind, ExprKind::Unary { op: UnaryOp::Deref, .. })
}

struct Native<'a, F> {
    resolve: &'a F,
    candidates: Vec<(String, BitLane, WordInterpretation)>,
    rejected: Vec<(String, Rejection)>,
}

impl<F: Fn(&TypeHint) -> Option<(BitLane, WordInterpretation)>> Native<'_, F> {
    fn reject(&mut self, name: &str, reason: Rejection) {
        if !self.rejected.iter().any(|(other, _)| other == name) {
            self.rejected.push((name.to_owned(), reason));
        }
    }

    /// Reject every candidate `test` reports, with `reason`.
    fn reject_where(&mut self, reason: Rejection, test: impl Fn(&str) -> bool) {
        let names: Vec<String> = self
            .candidates
            .iter()
            .map(|(name, ..)| name.clone())
            .filter(|name| test(name))
            .collect();
        for name in names {
            self.reject(&name, reason);
        }
    }

    fn declare(&mut self, name: &str, hint: Option<&TypeHint>) {
        let resolved = hint
            .filter(|hint| {
                hint.binding == TypeBinding::Checked
                    && hint.access == Some(StorageAccess::Ordinary)
            })
            .and_then(|hint| (self.resolve)(hint));
        let known = self.candidates.iter().find(|(other, ..)| other == name).cloned();
        match (resolved, known) {
            (Some((lane, interpretation)), None) => {
                if !self.rejected.iter().any(|(other, _)| other == name) {
                    self.candidates.push((name.to_owned(), lane, interpretation));
                }
            }
            (Some((lane, interpretation)), Some((_, known_lane, known_interpretation)))
                if lane == known_lane && interpretation == known_interpretation => {}
            (_, Some(_)) => self.reject(name, Rejection::ShadowedBinding),
            // A non-candidate binding of this name: a later candidate of the
            // same name would share one name-level proof with it.
            (None, None) => self.reject(name, Rejection::ShadowedBinding),
        }
    }

    fn mentions(&self, statement: &Statement) -> Vec<String> {
        let mut found = Vec::new();
        statement.walk_exprs(&mut |value| {
            for (name, ..) in &self.candidates {
                if contains(value, name) && !found.contains(name) {
                    found.push(name.clone());
                }
            }
        });
        found
    }

    fn expression(&mut self, value: &Expression) {
        let candidates: Vec<String> = self.candidates.iter().map(|(name, ..)| name.clone()).collect();
        let mut escapes: Vec<(String, Rejection)> = Vec::new();
        let mut bodies: Vec<Statement> = Vec::new();
        value.walk_exprs(&mut |node| {
            let mut escape = |scope: &Expression, reason: Rejection| {
                for name in &candidates {
                    if contains(scope, name) {
                        escapes.push((name.clone(), reason));
                    }
                }
            };
            match &node.kind {
                ExprKind::Assign { target, .. } if writes_through_pattern(target) => {
                    escape(target, Rejection::IncompatibleProducer)
                }
                ExprKind::RefOf(_) | ExprKind::Unary { op: UnaryOp::AddrOf, .. } => {
                    escape(node, Rejection::AddressOrReferenceEscape)
                }
                ExprKind::Unary {
                    op: UnaryOp::PostInc | UnaryOp::PreInc | UnaryOp::PostDec | UnaryOp::PreDec,
                    ..
                } => escape(node, Rejection::IncompatibleProducer),
                ExprKind::Call { args, .. } => {
                    for arg in args.iter().filter(|arg| arg.by_ref) {
                        escape(&arg.value, Rejection::AddressOrReferenceEscape);
                    }
                }
                ExprKind::Lambda { body, .. } => {
                    escape(node, Rejection::Capture);
                    match body {
                        vybe_ast::LambdaBody::Expr(value) => escape(value, Rejection::Capture),
                        vybe_ast::LambdaBody::Block(statements) => bodies.extend(statements.iter().cloned()),
                    }
                }
                ExprKind::FunctionExpr(statement) => bodies.push((**statement).clone()),
                ExprKind::Walrus { .. }
                | ExprKind::Destructure(_)
                | ExprKind::Comprehension { .. }
                | ExprKind::Atomic(_)
                | ExprKind::Async(_)
                | ExprKind::Yield(_)
                | ExprKind::YieldFrom(_)
                | ExprKind::Await(_) => escape(node, Rejection::UnsupportedExpression),
                _ => {}
            }
        });
        for statement in &bodies {
            for name in self.mentions(statement) {
                escapes.push((name, Rejection::Capture));
            }
        }
        for (name, reason) in escapes {
            self.reject(&name, reason);
        }
    }

    /// An expression whose value is not observed: a step of a bare name
    /// writes that name with a word of its own lane.
    fn discarded(&mut self, value: &Expression) {
        match &value.kind {
            ExprKind::Unary {
                op: UnaryOp::PostInc | UnaryOp::PreInc | UnaryOp::PostDec | UnaryOp::PreDec,
                expr,
            } if matches!(&expr.kind, ExprKind::Ident(_)) => {}
            ExprKind::Sequence(parts) => {
                for part in parts {
                    self.discarded(part);
                }
            }
            _ => self.expression(value),
        }
    }

    fn pattern(&mut self, pattern: &BindingPattern) {
        self.reject_where(Rejection::ShadowedBinding, |name| pattern_contains(pattern, name));
    }

    fn body(&mut self, body: &[Statement]) {
        for statement in body {
            self.statement(statement);
        }
    }

    fn statement(&mut self, statement: &Statement) {
        match &statement.kind {
            StmtKind::VarDecl { declarations, kind } => {
                for declaration in declarations {
                    if let Some(value) = &declaration.init {
                        self.expression(value);
                    }
                    match &declaration.pattern {
                        BindingPattern::Ident(name) => {
                            let ordinary = !matches!(kind, VarDeclKind::Static | VarDeclKind::FunctionScoped)
                                && declaration.array_bounds.is_none()
                                && !declaration.with_events;
                            self.declare(name, declaration.type_hint.as_ref().filter(|_| ordinary));
                        }
                        pattern => self.pattern(pattern),
                    }
                }
            }
            StmtKind::Expr(value) => self.discarded(value),
            StmtKind::Return(Some(value)) => self.expression(value),
            StmtKind::Throw { expr, cause } => {
                for value in [expr, cause].into_iter().flatten() {
                    self.expression(value);
                }
            }
            StmtKind::Assign { targets, value, by_ref } => {
                self.expression(value);
                let whole = targets.len() == 1
                    && !*by_ref
                    && matches!(&targets[0].kind, ExprKind::Ident(_));
                if !whole {
                    self.reject_where(Rejection::IncompatibleProducer, |name| {
                        targets.iter().any(|target| {
                            matches!(&target.kind, ExprKind::Ident(other) if other == name)
                                || (writes_through_pattern(target) && contains(target, name))
                        })
                    });
                }
                for target in targets {
                    self.expression(target);
                }
            }
            StmtKind::Return(None)
            | StmtKind::Break(_)
            | StmtKind::Continue(_)
            | StmtKind::Empty => {}
            StmtKind::Block(body) => self.body(body),
            StmtKind::Labeled { body, .. } => self.statement(body),
            StmtKind::If { cond, then_body, elifs, else_body } => {
                self.expression(cond);
                self.body(then_body);
                for (cond, body) in elifs {
                    self.expression(cond);
                    self.body(body);
                }
                if let Some(body) = else_body {
                    self.body(body);
                }
            }
            StmtKind::While { cond, body, else_body: None }
            | StmtKind::DoWhile { body, cond, until: false } => {
                self.expression(cond);
                self.body(body);
            }
            StmtKind::For { init, cond, update, body } => {
                if let Some(init) = init {
                    self.statement(init);
                }
                if let Some(cond) = cond {
                    self.expression(cond);
                }
                if let Some(update) = update {
                    self.discarded(update);
                }
                self.body(body);
            }
            StmtKind::ForIn { var, key, iter, body, else_body: None, .. } => {
                self.reject_where(Rejection::ShadowedBinding, |name| {
                    var == name || key.as_deref() == Some(name)
                });
                self.expression(iter);
                self.body(body);
            }
            StmtKind::Switch { expr, cases, default } => {
                self.expression(expr);
                for case in cases {
                    for condition in &case.conditions {
                        match condition {
                            vybe_ast::CaseCondition::Value(value)
                            | vybe_ast::CaseCondition::Comparison { expr: value, .. } => {
                                self.expression(value)
                            }
                            vybe_ast::CaseCondition::Range { from, to } => {
                                self.expression(from);
                                self.expression(to);
                            }
                        }
                    }
                    self.body(&case.body);
                }
                if let Some(body) = default {
                    self.body(body);
                }
            }
            // Statements outside the modeled set may access a binding by a
            // path this analysis does not see; any mention rejects it.
            _ => {
                for name in self.mentions(statement) {
                    self.reject(&name, Rejection::UnsupportedControlFlow);
                }
            }
        }
    }
}
