//! Close W64 intermediates and proven local storage inside a narrow boundary.
use std::collections::HashMap;
use vybe_ast::{
    BinOp, BindingPattern, BitLane, CaseCondition, DirectiveScope, Directives, ExprKind,
    Expression, Literal, Param, PassBy, Statement, StmtKind, StorageAccess, TypeBinding,
    VarDeclKind, WideIntegerTransport,
};

pub(crate) struct IntegerType {
    pub bits: u8,
    pub cast: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Width {
    Narrow,
    Wide,
}

fn cast(value: Expression, type_name: String) -> Expression {
    let span = value.span;
    Expression::with_span(
        ExprKind::Cast {
            expr: Box::new(value),
            type_name,
        },
        span,
    )
}

fn closed(
    value: &Expression,
    params: &HashMap<&str, IntegerType>,
    resolve: &impl Fn(&str) -> Option<IntegerType>,
) -> Option<(Expression, Width, bool)> {
    match &value.kind {
        ExprKind::Ident(name) => {
            let param = params.get(name.as_str())?;
            Some((
                cast(value.clone(), param.cast.clone()),
                Width::Narrow,
                false,
            ))
        }
        ExprKind::Lit(Literal::Int(word)) => Some((
            value.clone(),
            if (i64::from(i32::MIN)..=i64::from(u32::MAX)).contains(word) {
                Width::Narrow
            } else {
                Width::Wide
            },
            false,
        )),
        ExprKind::Cast { expr, type_name } => {
            let ty = resolve(type_name)?;
            let (inner, width, wide) = closed(expr, params, resolve)?;
            match ty.bits {
                // Typed Integer operations already own signedness/width. A
                // 64-bit integer cast changes no bits in their native lane.
                64 => Some((inner, Width::Wide, wide)),
                // A bare large Int literal is not an exact-word expression for
                // the shared narrowing-cast consumer. Keep its old path until
                // that boundary has an independent exactness proof.
                1..=32 if width == Width::Wide && matches!(inner.kind, ExprKind::Lit(_)) => None,
                1..=32 => Some((cast(inner, ty.cast), Width::Narrow, wide)),
                _ => None,
            }
        }
        ExprKind::Binary {
            op: BinOp::Integer(operation, lane),
            left,
            right,
        } => {
            let (left, left_width, left_wide) = closed(left, params, resolve)?;
            let (right, right_width, right_wide) = closed(right, params, resolve)?;
            let width = match lane {
                BitLane::W64 if !operation.is_comparison() => Width::Wide,
                BitLane::W32 if left_width == Width::Narrow && right_width == Width::Narrow => {
                    Width::Narrow
                }
                _ => return None,
            };
            Some((
                Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::Integer(*operation, *lane),
                        left: Box::new(left),
                        right: Box::new(right),
                    },
                    value.span,
                ),
                width,
                left_wide || right_wide || *lane == BitLane::W64,
            ))
        }
        _ => None,
    }
}

struct LocalRegion<'a, R> {
    declared: HashMap<String, IntegerType>,
    resolve: &'a R,
    wide_storage: bool,
    wide_switch: bool,
}

impl<R: Fn(&str) -> Option<IntegerType>> LocalRegion<'_, R> {
    fn exact_carrier(&self, value: &Expression) -> bool {
        match &value.kind {
            ExprKind::Ident(name) => self.declared.get(name).is_some_and(|ty| ty.bits == 64),
            ExprKind::Cast { type_name, .. } => {
                (self.resolve)(type_name).is_some_and(|ty| ty.bits == 64)
            }
            ExprKind::Binary {
                op: BinOp::Integer(op, BitLane::W64),
                ..
            } => !op.is_comparison(),
            _ => false,
        }
    }

    fn value(&self, value: &Expression) -> Option<Width> {
        match &value.kind {
            ExprKind::Ident(name) => Some(if self.declared.get(name)?.bits == 64 {
                Width::Wide
            } else {
                Width::Narrow
            }),
            ExprKind::Lit(Literal::Int(word)) => Some(
                if (i64::from(i32::MIN)..=i64::from(u32::MAX)).contains(word) {
                    Width::Narrow
                } else {
                    Width::Wide
                },
            ),
            ExprKind::Cast { expr, type_name } => {
                let width = self.value(expr)?;
                Some(match (self.resolve)(type_name)?.bits {
                    64 => Width::Wide,
                    1..=32 if width == Width::Narrow || self.exact_carrier(expr) => Width::Narrow,
                    _ => return None,
                })
            }
            ExprKind::Binary {
                op: BinOp::Integer(op, lane),
                left,
                right,
            } => {
                let left = self.value(left)?;
                let right = self.value(right)?;
                match lane {
                    BitLane::W64 => Some(if op.is_comparison() {
                        Width::Narrow
                    } else {
                        Width::Wide
                    }),
                    BitLane::W32 if left == Width::Narrow && right == Width::Narrow => {
                        Some(Width::Narrow)
                    }
                    _ => None,
                }
            }
            ExprKind::Binary {
                op:
                    BinOp::Eq
                    | BinOp::NotEq
                    | BinOp::Lt
                    | BinOp::LtEq
                    | BinOp::Gt
                    | BinOp::GtEq
                    | BinOp::And
                    | BinOp::Or,
                left,
                right,
            } if self.value(left)? == Width::Narrow && self.value(right)? == Width::Narrow => {
                Some(Width::Narrow)
            }
            // In particular, a declared cast around a call is NOT an exact
            // return-transport contract. Reject calls, captures and aliasing.
            _ => None,
        }
    }

    fn statements(
        &mut self,
        body: &mut [Statement],
        result: &IntegerType,
        top: bool,
    ) -> Option<()> {
        for statement in body {
            match &mut statement.kind {
                StmtKind::VarDecl { declarations, kind } if top && *kind != VarDeclKind::Static => {
                    for declaration in declarations {
                        let BindingPattern::Ident(name) = &declaration.pattern else {
                            return None;
                        };
                        let hint = declaration.type_hint.as_ref()?;
                        if declaration.array_bounds.is_some()
                            || declaration.with_events
                            || hint.binding == TypeBinding::Descriptive
                            || hint.access != Some(StorageAccess::Ordinary)
                            || self.declared.contains_key(name)
                        {
                            return None;
                        }
                        let ty = (self.resolve)(hint.spelling())?;
                        if !matches!(ty.bits, 1..=32 | 64) {
                            return None;
                        }
                        let init = declaration.init.as_mut()?;
                        if self.value(init)? == Width::Wide
                            && ty.bits != 64
                            && !self.exact_carrier(init)
                        {
                            return None;
                        }
                        *init = cast(init.clone(), ty.cast.clone());
                        self.wide_storage |= ty.bits == 64;
                        self.declared.insert(name.clone(), ty);
                    }
                }
                StmtKind::Expr(value) => {
                    let ExprKind::Assign { target, value } = &mut value.kind else {
                        return None;
                    };
                    let ExprKind::Ident(name) = &target.kind else {
                        return None;
                    };
                    let ty = self.declared.get(name)?;
                    if self.value(value)? == Width::Wide
                        && ty.bits != 64
                        && !self.exact_carrier(value)
                    {
                        return None;
                    }
                    **value = cast((**value).clone(), ty.cast.clone());
                }
                StmtKind::Switch {
                    expr,
                    cases,
                    default,
                } => {
                    if self.value(expr)? != Width::Wide || !self.exact_carrier(expr) {
                        return None;
                    }
                    self.wide_switch = true;
                    for case in cases {
                        for condition in &case.conditions {
                            let CaseCondition::Value(value) = condition else {
                                return None;
                            };
                            if self.value(value)? != Width::Wide || !self.exact_carrier(value) {
                                return None;
                            }
                        }
                        self.statements(&mut case.body, result, false)?;
                    }
                    if let Some(default) = default {
                        self.statements(default, result, false)?;
                    }
                }
                StmtKind::If {
                    cond,
                    then_body,
                    elifs,
                    else_body,
                } => {
                    if self.value(cond)? != Width::Narrow {
                        return None;
                    }
                    self.statements(then_body, result, false)?;
                    for (cond, body) in elifs {
                        if self.value(cond)? != Width::Narrow {
                            return None;
                        }
                        self.statements(body, result, false)?;
                    }
                    if let Some(body) = else_body {
                        self.statements(body, result, false)?;
                    }
                }
                StmtKind::Block(body) => self.statements(body, result, false)?,
                StmtKind::Return(Some(value)) => {
                    if self.value(value)? != Width::Narrow {
                        return None;
                    }
                    *value = cast(value.clone(), result.cast.clone());
                }
                StmtKind::Break(vybe_ast::BreakTarget::Implicit) => {}
                // Existing directives must retain their scope and policy;
                // unknown control flow, storage and nested functions fail closed.
                _ => return None,
            }
        }
        Some(())
    }
}

fn local_storage_region(
    params: &[Param],
    result: &IntegerType,
    body: &[Statement],
    resolve: &impl Fn(&str) -> Option<IntegerType>,
) -> Option<Vec<Statement>> {
    // Most functions have no candidate storage. Do not clone their lowered
    // bodies just to discover that the transport policy cannot apply.
    if !body.iter().any(|statement| match &statement.kind {
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|declaration| {
            declaration
                .type_hint
                .as_ref()
                .is_some_and(|hint| resolve(hint.spelling()).is_some_and(|ty| ty.bits == 64))
        }),
        _ => false,
    }) {
        return None;
    }
    let mut region = LocalRegion {
        declared: HashMap::new(),
        resolve,
        wide_storage: false,
        wide_switch: false,
    };
    for param in params {
        let hint = param.type_hint.as_ref()?;
        if param.pass_by != PassBy::Value
            || param.is_rest
            || param.is_kwargs
            || param.is_optional
            || param.is_nullable
            || param.default.is_some()
            || hint.binding == TypeBinding::Descriptive
            || hint.access != Some(StorageAccess::Ordinary)
        {
            return None;
        }
        let ty = resolve(hint.spelling()).filter(|ty| (1..=32).contains(&ty.bits))?;
        if region.declared.insert(param.name.clone(), ty).is_some() {
            return None;
        }
    }
    let mut candidate = body.to_vec();
    region.statements(&mut candidate, result, true)?;
    (region.wide_storage && region.wide_switch).then_some(candidate)
}

pub(crate) fn exact_directive() -> Statement {
    Statement::new(StmtKind::Directive {
        set: Directives {
            wide_integer_transport: Some(WideIntegerTransport::Exact),
            ..Default::default()
        },
        scope: DirectiveScope::Block,
    })
}

/// A closed record-selector boundary does not change the enclosing function ABI.
/// The walker certifies the producer identity and installs its helper-return
/// directive only after all comparison/result consumers have passed this proof.
pub(crate) fn normalize_record_switch(
    statement: &mut Statement,
    certify: impl FnOnce(&mut Expression) -> bool,
) -> bool {
    fn constant(value: &Expression) -> bool {
        match &value.kind {
            ExprKind::Lit(Literal::Int(_) | Literal::BigInt(_)) => true,
            ExprKind::Cast { expr, type_name }
                if matches!(type_name.as_str(), "int64" | "uint64") =>
            {
                constant(expr)
            }
            _ => false,
        }
    }
    fn narrow_body(body: &[Statement]) -> bool {
        body.iter().all(|statement| match &statement.kind {
            StmtKind::Return(Some(value)) => matches!(&value.kind,
                ExprKind::Lit(Literal::Int(word)) if i32::try_from(*word).is_ok()),
            StmtKind::Break(vybe_ast::BreakTarget::Implicit) => true,
            StmtKind::Block(body) => narrow_body(body),
            _ => false,
        })
    }
    let StmtKind::Switch {
        expr,
        cases,
        default,
    } = &statement.kind
    else {
        return false;
    };
    if !matches!(&expr.kind, ExprKind::Cast { type_name, .. }
        if matches!(type_name.as_str(), "int64" | "uint64"))
        || !cases.iter().all(|case| {
            case.conditions.iter().all(|condition| {
                matches!(condition,
                CaseCondition::Value(value) if constant(value))
            }) && narrow_body(&case.body)
        })
        || !default.as_ref().is_none_or(|body| narrow_body(body))
    {
        return false;
    }
    let mut candidate = statement.clone();
    let StmtKind::Switch { expr, .. } = &mut candidate.kind else {
        unreachable!();
    };
    if !certify(expr) {
        return false;
    }
    let StmtKind::Switch { cases, .. } = &mut candidate.kind else {
        unreachable!();
    };
    // Case values cross the boxed switch operand boundary independently of
    // the certified producer. Carry their bits explicitly, before any Number
    // adaptation can round a positive word above 2^53.
    for case in cases {
        for condition in &mut case.conditions {
            let CaseCondition::Value(value) = condition else { unreachable!(); };
            value.walk_exprs_mut(&mut |value| {
                if let ExprKind::Lit(Literal::Int(word)) = value.kind {
                    value.kind = ExprKind::Lit(Literal::BigInt(word));
                }
            });
        }
    }
    statement.kind = StmtKind::Block(vec![exact_directive(), candidate]);
    true
}

/// Reject uncertainty before mutating the body. Only initialized ordinary
/// scalar locals may add storage; no wide ABI input or result is admitted.
pub(crate) fn normalize(
    params: &[Param],
    return_type: Option<&str>,
    body: &mut Vec<Statement>,
    resolve: impl Fn(&str) -> Option<IntegerType>,
) -> bool {
    let Some(result) = return_type
        .and_then(&resolve)
        .filter(|ty| (1..=32).contains(&ty.bits))
    else {
        return false;
    };
    if let Some(mut candidate) = local_storage_region(params, &result, body, &resolve) {
        candidate.insert(
            0,
            Statement::new(StmtKind::Directive {
                set: Directives {
                    wide_integer_transport: Some(WideIntegerTransport::Exact),
                    ..Default::default()
                },
                scope: DirectiveScope::Block,
            }),
        );
        *body = candidate;
        return true;
    }
    let [statement] = body.as_slice() else {
        return false;
    };
    let StmtKind::Return(Some(value)) = &statement.kind else {
        return false;
    };
    let mut declared = HashMap::new();
    for param in params {
        let Some(hint) = &param.type_hint else {
            return false;
        };
        if param.pass_by != PassBy::Value
            || param.is_rest
            || param.is_kwargs
            || param.is_optional
            || param.is_nullable
            || param.default.is_some()
            || hint.binding == TypeBinding::Descriptive
            || hint.access != Some(StorageAccess::Ordinary)
        {
            return false;
        }
        let Some(ty) = resolve(hint.spelling()).filter(|ty| (1..=32).contains(&ty.bits)) else {
            return false;
        };
        if declared.insert(param.name.as_str(), ty).is_some() {
            return false;
        }
    }
    let Some((value, _, true)) = closed(value, &declared, &resolve) else {
        return false;
    };
    body[0].kind = StmtKind::Return(Some(cast(value, result.cast)));
    body.insert(
        0,
        Statement::new(StmtKind::Directive {
            set: Directives {
                wide_integer_transport: Some(WideIntegerTransport::Exact),
                ..Default::default()
            },
            scope: DirectiveScope::Block,
        }),
    );
    true
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{Argument, IntOp, Span, UnaryOp};

    #[test]
    fn certified_record_switch_constants_cross_boxed_boundary_exactly() {
        fn find_switch(body: &[Statement]) -> Option<&Statement> {
            body.iter().find_map(|statement| match &statement.kind {
                StmtKind::Switch { .. } => Some(statement),
                StmtKind::Block(body) => find_switch(body),
                _ => None,
            })
        }
        let source = "struct R { unsigned long long word; }; int f(struct R *p) {
            switch(p->word) {
            case 0x1fffffffffffffULL: return 1;
            case 0x20000000000000ULL: return 2;
            case 0x20000000000001ULL: return 3;
            case 0x20000000000002ULL: return 4;
            case 0xfffffffffffffffeULL: return 5;
            case 0xffffffffffffffffULL: return 6;
            default: return 7; }}";
        let mut baseline = None;
        for parser in [crate::parse_with_peg, crate::parse_optimized] {
            let module = parser(source).unwrap();
            assert_eq!(module.directives.wide_integer_transport, None);
            let function = module.body.iter().find(|statement|
                matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "f")).unwrap();
            let StmtKind::FunctionDecl { body, .. } = &function.kind else { unreachable!(); };
            let StmtKind::Switch { cases, .. } = &find_switch(body).unwrap().kind else { unreachable!(); };
            assert_eq!(cases.len(), 7, "six value cases and ordered default");
            assert!(cases.last().unwrap().conditions.is_empty());
            let mut constants = Vec::new();
            for case in &cases[..6] {
                let [CaseCondition::Value(value)] = case.conditions.as_slice() else {
                    panic!("one value per case");
                };
                value.walk_exprs(&mut |value| {
                    if let ExprKind::Lit(Literal::BigInt(word)) = &value.kind {
                        constants.push(*word);
                    }
                    assert!(!matches!(value.kind, ExprKind::Lit(Literal::Int(_))),
                        "certified case words must cross the boxed boundary exactly");
                });
            }
            assert_eq!(constants, vec![0x1fffffffffffff, 0x20000000000000,
                0x20000000000001, 0x20000000000002, -2, -1]);
            let ast = format!("{module:?}");
            if let Some(previous) = &baseline { assert_eq!(&ast, previous); }
            baseline = Some(ast);
        }

        let source = "unsigned long long unknown(void); int f(void) {
            switch(unknown()) { case 0x20000000000001ULL:return 3; default:return 7; }}";
        let module = crate::parse_with_peg(source).unwrap();
        let body = module.body.iter().find_map(|statement| match &statement.kind {
            StmtKind::FunctionDecl { name, body, .. } if name == "f" => Some(body),
            _ => None,
        }).unwrap();
        let StmtKind::Switch { cases, .. } = &find_switch(body).unwrap().kind else { unreachable!(); };
        let CaseCondition::Value(value) = &cases[0].conditions[0] else { unreachable!(); };
        let mut original = Vec::new();
        value.walk_exprs(&mut |value| {
            assert!(!matches!(value.kind, ExprKind::Lit(Literal::BigInt(_))),
                "unknown producers must not acquire the record proof");
            if let ExprKind::Lit(Literal::Int(word)) = &value.kind { original.push(*word); }
        });
        assert_eq!(original, vec![0x20000000000001]);
    }

    #[test]
    fn record_switch_requires_certified_producer_and_closed_narrow_consumers() {
        let source = "struct R { unsigned long long word; }; int f(struct R *p) {
            switch(p->word) { case 0xffffffffffffffffULL:return 33; default:return 44; }}";
        let module = crate::parse(source).unwrap();
        assert_eq!(module.directives.wide_integer_transport, None);
        let body = module
            .body
            .iter()
            .find_map(|statement| match &statement.kind {
                StmtKind::FunctionDecl { name, body, .. } if name == "f" => Some(body),
                _ => None,
            })
            .unwrap();
        assert!(body.iter().any(|statement| matches!(&statement.kind,
            StmtKind::Block(body) if matches!(&body[0].kind,
                StmtKind::Directive { set, .. } if set.wide_integer_transport == Some(WideIntegerTransport::Exact)))),
            "certified record selector did not close: {body:#?}");
        for source in [
            "unsigned long long unknown(void); int f(void) {switch(unknown()) {case 1:return 2; default:return 3;}}",
            "int f(unsigned long long p) {switch(p) {case 1:return 2; default:return 3;}}",
            "struct R { unsigned long long word; }; unsigned long long f(struct R *p) {switch(p->word) {case 1:return 2; default:return 3;}}",
            "struct R { unsigned long long word; }; int other(void); int f(struct R *p) {switch(p->word) {case 1:return other(); default:return 3;}}",
            "struct R { volatile unsigned long long word; }; int f(struct R *p) {switch(p->word) {case 1:return 2; default:return 3;}}",
        ] {
            let module = crate::parse(source).unwrap();
            let mut exact = 0;
            for statement in &module.body {
                statement.walk_exprs(&mut |expression| {
                    if let ExprKind::Lambda { body: vybe_ast::LambdaBody::Block(body), .. } = &expression.kind {
                        exact += usize::from(body.iter().any(|statement| matches!(&statement.kind,
                            StmtKind::Directive { set, .. } if set.wide_integer_transport == Some(WideIntegerTransport::Exact))));
                    }
                });
            }
            assert_eq!(
                exact, 0,
                "uncertified return/ABI/access must not acquire helper proof: {source}"
            );
        }
    }

    fn resolve(name: &str) -> Option<IntegerType> {
        let (bits, emitted) = match name {
            "fixed_t" | "int" => (32, "int"),
            "unsigned" | "uint32" => (32, "uint32"),
            "int64_t" | "int64" => (64, "int64"),
            "unsigned long long" | "uint64" => (64, "uint64"),
            _ => return None,
        };
        Some(IntegerType {
            bits,
            cast: emitted.into(),
        })
    }

    fn params() -> Vec<Param> {
        ["a", "b"]
            .into_iter()
            .map(|name| {
                let mut hint: vybe_ast::TypeHint = "fixed_t".into();
                hint.access = Some(StorageAccess::Ordinary);
                Param {
                    name: name.into(),
                    type_hint: Some(hint),
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }
            })
            .collect()
    }

    fn binary(op: IntOp, lane: BitLane, left: Expression, right: Expression) -> Expression {
        Expression::new(ExprKind::Binary {
            op: BinOp::Integer(op, lane),
            left: Box::new(left),
            right: Box::new(right),
        })
    }

    fn fixed_product() -> Expression {
        binary(
            IntOp::ShrS,
            BitLane::W64,
            binary(
                IntOp::Mul,
                BitLane::W64,
                cast(Expression::ident("a"), "int64_t".into()),
                cast(Expression::ident("b"), "int64_t".into()),
            ),
            Expression::int(16),
        )
    }

    fn local_switch_body() -> Vec<Statement> {
        let mut hint = vybe_ast::TypeHint::checked("uint64");
        hint.access = Some(StorageAccess::Ordinary);
        vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![vybe_ast::VarDeclarator {
                    pattern: BindingPattern::Ident("word".into()),
                    type_hint: Some(hint),
                    init: Some(cast(Expression::int(-2), "uint64".into())),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Var,
            }),
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("word")),
                value: Box::new(cast(Expression::int(-1), "uint64".into())),
            }))),
            Statement::new(StmtKind::Switch {
                expr: cast(Expression::ident("word"), "uint64".into()),
                cases: vec![vybe_ast::SwitchCase {
                    conditions: vec![CaseCondition::Value(cast(
                        Expression::int(-1),
                        "uint64".into(),
                    ))],
                    body: vec![Statement::new(StmtKind::Return(Some(Expression::int(33))))],
                }],
                default: Some(vec![Statement::new(StmtKind::Return(Some(
                    Expression::int(44),
                )))]),
            }),
        ]
    }

    #[test]
    fn local_w64_initialization_assignment_and_switch_share_one_exact_region() {
        let mut body = local_switch_body();
        let original_span = body[0].span;
        assert!(normalize(&[], Some("int"), &mut body, resolve));
        assert!(
            matches!(&body[0].kind, StmtKind::Directive { set, scope: DirectiveScope::Block }
            if set.wide_integer_transport == Some(WideIntegerTransport::Exact))
        );
        assert_eq!(format!("{:?}", body[1].span), format!("{original_span:?}"));
        let StmtKind::VarDecl { declarations, .. } = &body[1].kind else {
            panic!("local");
        };
        assert!(matches!(&declarations[0].init.as_ref().unwrap().kind,
            ExprKind::Cast { type_name, .. } if type_name == "uint64"));
        let StmtKind::Expr(value) = &body[2].kind else {
            panic!("assignment");
        };
        let ExprKind::Assign { value, .. } = &value.kind else {
            panic!("assignment");
        };
        assert!(matches!(&value.kind, ExprKind::Cast { type_name, .. } if type_name == "uint64"));
        let before = format!("{body:?}");
        assert!(!normalize(&[], Some("int"), &mut body, resolve));
        assert_eq!(before, format!("{body:?}"));
    }

    #[test]
    fn local_w64_region_rejects_unknown_call_even_under_declared_cast() {
        for position in [0, 1] {
            let mut body = local_switch_body();
            let unknown = cast(
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("unknown")),
                    args: vec![],
                    optional: false,
                }),
                "uint64".into(),
            );
            if position == 0 {
                let StmtKind::VarDecl { declarations, .. } = &mut body[0].kind else {
                    unreachable!()
                };
                declarations[0].init = Some(unknown);
            } else {
                let StmtKind::Expr(value) = &mut body[1].kind else {
                    unreachable!()
                };
                let ExprKind::Assign { value, .. } = &mut value.kind else {
                    unreachable!()
                };
                **value = unknown;
            }
            let before = format!("{body:?}");
            assert!(!normalize(&[], Some("int"), &mut body, resolve));
            assert_eq!(before, format!("{body:?}"));
        }
    }

    #[test]
    fn local_w64_region_preserves_directives_and_rejects_unproven_storage() {
        let mut variants = Vec::new();
        let mut body = local_switch_body();
        body.insert(
            0,
            Statement::new(StmtKind::Directive {
                set: Directives {
                    wide_integer_transport: Some(WideIntegerTransport::Number),
                    ..Default::default()
                },
                scope: DirectiveScope::Block,
            }),
        );
        variants.push(body);
        for access in [
            None,
            Some(StorageAccess::Volatile),
            Some(StorageAccess::Atomic),
        ] {
            let mut body = local_switch_body();
            let StmtKind::VarDecl { declarations, .. } = &mut body[0].kind else {
                unreachable!()
            };
            declarations[0].type_hint.as_mut().unwrap().access = access;
            variants.push(body);
        }
        let mut body = local_switch_body();
        let StmtKind::Expr(value) = &mut body[1].kind else {
            unreachable!()
        };
        let ExprKind::Assign { target, .. } = &mut value.kind else {
            unreachable!()
        };
        **target = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("record")),
            field: "word".into(),
            null_safe: false,
        });
        variants.push(body);
        let mut body = local_switch_body();
        let StmtKind::Switch { cases, .. } = &mut body[2].kind else {
            unreachable!()
        };
        cases[0].conditions[0] = CaseCondition::Range {
            from: cast(Expression::int(-2), "uint64".into()),
            to: cast(Expression::int(-1), "uint64".into()),
        };
        variants.push(body);
        for mut body in variants {
            let before = format!("{body:?}");
            assert!(!normalize(&[], Some("int"), &mut body, resolve));
            assert_eq!(before, format!("{body:?}"));
        }
        let mut wide_params = params();
        wide_params[0]
            .type_hint
            .as_mut()
            .unwrap()
            .set_spelling("uint64");
        assert!(!normalize(
            &wide_params,
            Some("int"),
            &mut local_switch_body(),
            resolve
        ));
        assert!(!normalize(
            &[],
            Some("uint64"),
            &mut local_switch_body(),
            resolve
        ));
    }

    #[test]
    fn source_local_w64_switch_activates_without_module_or_abi_policy() {
        let module = crate::parse(
            "int f(void) { unsigned long long word = 0xfffffffffffffffeULL;
            word = 0xffffffffffffffffULL;
            switch (word) { case 0xfffffffffffffffeULL: return 22;
                case 0xffffffffffffffffULL: return 33; default: return 44; } }",
        )
        .unwrap();
        assert_eq!(module.directives.wide_integer_transport, None);
        let body = module
            .body
            .iter()
            .find_map(|statement| match &statement.kind {
                StmtKind::FunctionDecl { name, body, .. } if name == "f" => Some(body),
                _ => None,
            })
            .unwrap();
        assert!(matches!(&body[0].kind, StmtKind::Directive { set, .. }
            if set.wide_integer_transport == Some(WideIntegerTransport::Exact)));
    }

    #[test]
    fn closed_wide_arithmetic_gets_local_policy_and_explicit_narrow_boundary() {
        let span = Span {
            start_line: 4,
            start_col: 5,
            end_line: 4,
            end_col: 40,
        };
        let mut value = fixed_product();
        value.span = span;
        let mut body = vec![Statement::with_span(StmtKind::Return(Some(value)), span)];
        assert!(normalize(&params(), Some("fixed_t"), &mut body, resolve));
        assert!(
            matches!(&body[0].kind, StmtKind::Directive { set, scope: DirectiveScope::Block }
            if set.wide_integer_transport == Some(WideIntegerTransport::Exact))
        );
        let StmtKind::Return(Some(value)) = &body[1].kind else {
            panic!("return");
        };
        assert_eq!(format!("{:?}", value.span), format!("{span:?}"));
        assert_eq!(format!("{:?}", body[1].span), format!("{span:?}"));
        assert!(matches!(&value.kind, ExprKind::Cast { type_name, .. } if type_name == "int"));
        let dump = format!("{value:?}");
        assert!(!dump.contains("int64_t"));
        assert!(dump.contains("W64"));
        let before = format!("{body:?}");
        assert!(!normalize(&params(), Some("fixed_t"), &mut body, resolve));
        assert_eq!(before, format!("{body:?}"));
    }

    #[test]
    fn uncertain_inputs_and_storage_leave_body_unchanged() {
        for value in [
            Expression::ident("global"),
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("f")),
                args: vec![Argument::positional(Expression::ident("a"))],
                optional: false,
            }),
            Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("a")),
                field: "x".into(),
                null_safe: false,
            }),
            Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("a")),
                value: Box::new(Expression::int(1)),
            }),
        ] {
            let value = binary(IntOp::Add, BitLane::W64, fixed_product(), value);
            let mut body = vec![Statement::new(StmtKind::Return(Some(value)))];
            let before = format!("{body:?}");
            assert!(!normalize(&params(), Some("fixed_t"), &mut body, resolve));
            assert_eq!(before, format!("{body:?}"));
        }
        for (param_type, return_type) in [
            ("int64_t", "fixed_t"),
            ("int*", "fixed_t"),
            ("record", "fixed_t"),
            ("fixed_t", "int64_t"),
        ] {
            let mut params = params();
            params[0]
                .type_hint
                .as_mut()
                .unwrap()
                .set_spelling(param_type);
            let mut body = vec![Statement::new(StmtKind::Return(Some(fixed_product())))];
            let before = format!("{body:?}");
            assert!(!normalize(&params, Some(return_type), &mut body, resolve));
            assert_eq!(before, format!("{body:?}"));
        }
    }

    #[test]
    fn no_w64_or_unnarrowed_wide_input_does_not_activate() {
        for value in [
            binary(
                IntOp::Add,
                BitLane::W32,
                Expression::ident("a"),
                Expression::int(2),
            ),
            binary(
                IntOp::Add,
                BitLane::W32,
                fixed_product(),
                Expression::int(2),
            ),
            binary(IntOp::Eq, BitLane::W64, fixed_product(), Expression::int(2)),
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(fixed_product()),
            }),
            binary(
                IntOp::Add,
                BitLane::W64,
                fixed_product(),
                cast(Expression::int(9_007_199_254_740_993), "int".into()),
            ),
        ] {
            let mut body = vec![Statement::new(StmtKind::Return(Some(value)))];
            let before = format!("{body:?}");
            assert!(!normalize(&params(), Some("fixed_t"), &mut body, resolve));
            assert_eq!(before, format!("{body:?}"));
        }
    }

    #[test]
    fn declarations_effectful_access_and_nonvalue_parameters_are_not_regions() {
        for access in [
            None,
            Some(StorageAccess::Volatile),
            Some(StorageAccess::Atomic),
        ] {
            let mut params = params();
            params[0].type_hint.as_mut().unwrap().access = access;
            let mut body = vec![Statement::new(StmtKind::Return(Some(fixed_product())))];
            let before = format!("{body:?}");
            assert!(!normalize(&params, Some("fixed_t"), &mut body, resolve));
            assert_eq!(before, format!("{body:?}"));
        }
        for pass_by in [PassBy::Alias, PassBy::Ref, PassBy::Out] {
            let mut params = params();
            params[0].pass_by = pass_by;
            let mut body = vec![Statement::new(StmtKind::Return(Some(fixed_product())))];
            let before = format!("{body:?}");
            assert!(!normalize(&params, Some("fixed_t"), &mut body, resolve));
            assert_eq!(before, format!("{body:?}"));
        }
        for source in [
            "int f(int a) { long long x=(long long)a; return (int)(x*x); }",
            "int f(int a) { int x=a; return (int)((long long)x*x); }",
            "int global; int f(int a) { return (int)((long long)a*global); }",
            "int other(int); int f(int a) { return (int)((long long)a*other(a)); }",
            "struct R { int x; }; int f(struct R r) { return (int)((long long)r.x*r.x); }",
            "long long f(int a) { return (long long)a*a; }",
            "int f(long long a) { return (int)(a*a); }",
            "int f(int a) { return (int)(-((long long)a*a)); }",
        ] {
            let module = crate::parse(source).unwrap();
            let body = module
                .body
                .iter()
                .find_map(|statement| {
                    if let StmtKind::FunctionDecl { name, body, .. } = &statement.kind {
                        (name == "f").then_some(body)
                    } else {
                        None
                    }
                })
                .unwrap();
            assert!(!body.iter().any(|statement| matches!(&statement.kind,
                StmtKind::Directive { set, .. } if set.wide_integer_transport == Some(WideIntegerTransport::Exact))), "{source}");
        }
    }

    #[test]
    fn source_fixed_product_activates_without_function_name_or_module_policy() {
        let source = "typedef int fixed_t; typedef long long int64_t;
            fixed_t unrelated(fixed_t a, fixed_t b) { return ((int64_t)a * (int64_t)b) >> 16; }
            int narrow(int a, int b) { return a+b; }";
        let module = crate::parse(source).unwrap();
        assert_eq!(module.directives.wide_integer_transport, None);
        let bodies = module
            .body
            .iter()
            .filter_map(|statement| {
                if let StmtKind::FunctionDecl { name, body, .. } = &statement.kind {
                    Some((name.as_str(), body))
                } else {
                    None
                }
            })
            .collect::<HashMap<_, _>>();
        assert!(matches!(
            bodies["unrelated"][0].kind,
            StmtKind::Directive { .. }
        ));
        assert!(!matches!(
            bodies["narrow"][0].kind,
            StmtKind::Directive { .. }
        ));
    }
}
