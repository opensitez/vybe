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
