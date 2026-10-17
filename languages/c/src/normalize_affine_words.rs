//! Guarded indexed 16-bit scatter using the sequential libc word-column entry.
use super::*;
use std::collections::BTreeMap;

#[derive(Clone, Debug, PartialEq, Eq)]
struct Affine {
    terms: BTreeMap<String, i64>,
    offset: i64,
}

impl Affine {
    fn constant(offset: i64) -> Self {
        Self {
            terms: BTreeMap::new(),
            offset,
        }
    }
    fn variable(name: &str) -> Self {
        Self {
            terms: [(name.into(), 1)].into(),
            offset: 0,
        }
    }
    fn scaled(mut self, scale: i64) -> Option<Self> {
        self.offset = self.offset.checked_mul(scale)?;
        for coefficient in self.terms.values_mut() {
            *coefficient = coefficient.checked_mul(scale)?;
        }
        self.terms.retain(|_, coefficient| *coefficient != 0);
        Some(self)
    }
    fn plus(mut self, other: Self) -> Option<Self> {
        self.offset = self.offset.checked_add(other.offset)?;
        for (name, coefficient) in other.terms {
            let entry = self.terms.entry(name).or_default();
            *entry = entry.checked_add(coefficient)?;
        }
        self.terms.retain(|_, coefficient| *coefficient != 0);
        Some(self)
    }
}

#[derive(Clone, Debug, PartialEq)]
enum Symbol {
    Address(Affine),
    Word(Affine),
    Boolean(bool),
    NumberType,
}

struct Evaluation {
    values: HashMap<String, Symbol>,
    writes: HashSet<String>,
    reads: Vec<Affine>,
    stores: Vec<(Affine, Affine)>,
}

impl Evaluation {
    fn address(&mut self, expression: &Expression) -> Option<Affine> {
        match self.evaluate(expression)? {
            Symbol::Address(value) => Some(value),
            _ => None,
        }
    }
    fn evaluate(&mut self, expression: &Expression) -> Option<Symbol> {
        match &expression.kind {
            ExprKind::Ident(name) => self.values.get(name).cloned(),
            ExprKind::Lit(Literal::Int(n)) => Some(Symbol::Address(Affine::constant(*n))),
            ExprKind::Lit(Literal::Str(s)) if s == "number" => Some(Symbol::NumberType),
            ExprKind::Assign { target, value } => {
                let name = ident(target)?.to_owned();
                let value = self.evaluate(value)?;
                self.writes.insert(name.clone());
                self.values.insert(name, value.clone());
                Some(value)
            }
            ExprKind::Sequence(values) => {
                let mut last = None;
                for value in values {
                    last = Some(self.evaluate(value)?);
                }
                last
            }
            ExprKind::IsType { expr, type_name } if type_name == "number" => {
                self.address(expr)?;
                Some(Symbol::Boolean(true))
            }
            ExprKind::Unary {
                op: UnaryOp::Typeof,
                expr,
            } => {
                self.address(expr)?;
                Some(Symbol::NumberType)
            }
            ExprKind::Binary {
                op: BinOp::Eq,
                left,
                right,
            } => match (self.evaluate(left)?, self.evaluate(right)?) {
                (Symbol::NumberType, Symbol::NumberType) => Some(Symbol::Boolean(true)),
                _ => None,
            },
            ExprKind::Ternary { cond, then, else_ } => match self.evaluate(cond)? {
                Symbol::Boolean(true) => self.evaluate(then),
                Symbol::Boolean(false) => self.evaluate(else_),
                _ => None,
            },
            ExprKind::Binary {
                op: BinOp::BitAnd | BinOp::Integer(IntOp::And, BitLane::W32),
                left,
                right,
            } => {
                let (value, mask) = match (&left.kind, &right.kind) {
                    (_, ExprKind::Lit(Literal::Int(mask))) => (left, *mask),
                    (ExprKind::Lit(Literal::Int(mask)), _) => (right, *mask),
                    _ => return None,
                };
                if !(i64::from(i32::MIN)..=i64::from(u32::MAX)).contains(&mask)
                    || mask & 0xffff != 0xffff
                {
                    return None;
                }
                // Word tracks the stored low 16 bits, not sign extension. Only
                // a proven word load can survive this bit-preserving mask.
                let value = self.evaluate(value)?;
                matches!(&value, Symbol::Word(_)).then_some(value)
            }
            ExprKind::Binary { op, left, right } => {
                let left = self.address(left)?;
                let right = self.address(right)?;
                let value = match op {
                    BinOp::Add | BinOp::Integer(IntOp::Add, BitLane::W32) => left.plus(right)?,
                    BinOp::Sub | BinOp::Integer(IntOp::Sub, BitLane::W32) => {
                        left.plus(right.scaled(-1)?)?
                    }
                    BinOp::Mul | BinOp::Integer(IntOp::Mul, BitLane::W32)
                        if right.terms.is_empty() =>
                    {
                        left.scaled(right.offset)?
                    }
                    BinOp::Mul | BinOp::Integer(IntOp::Mul, BitLane::W32)
                        if left.terms.is_empty() =>
                    {
                        right.scaled(left.offset)?
                    }
                    _ => return None,
                };
                Some(Symbol::Address(value))
            }
            ExprKind::Cast { expr, type_name } if !type_name.contains("volatile") => {
                let value = self.evaluate(expr)?;
                match &value {
                    Symbol::Address(_)
                        if type_name.contains('*')
                            || matches!(
                                type_name.as_str(),
                                "pointerCast" | "int" | "int32" | "integer"
                            ) =>
                    {
                        Some(value)
                    }
                    Symbol::Word(_)
                        if matches!(
                            type_name.as_str(),
                            "int"
                                | "integer"
                                | "int16"
                                | "uint16"
                                | "short"
                                | "short int"
                                | "signed short"
                                | "signed short int"
                                | "unsigned short"
                                | "unsigned short int"
                        ) =>
                    {
                        Some(value)
                    }
                    _ => None,
                }
            }
            ExprKind::Call {
                callee,
                args,
                optional: false,
            } if args
                .iter()
                .all(|arg| !arg.by_ref && !arg.spread && arg.name.is_none()) =>
            {
                match (ident(callee)?, args.as_slice()) {
                    ("__c_ptr_i32_load16_u" | "__c_ptr_i32_load16_s", [address]) => {
                        let address = self.address(&address.value)?;
                        self.reads.push(address.clone());
                        Some(Symbol::Word(address))
                    }
                    ("__c_ptr_i32_store16", [address, value]) => {
                        let address = self.address(&address.value)?;
                        let Symbol::Word(source) = self.evaluate(&value.value)? else {
                            return None;
                        };
                        self.stores.push((address, source.clone()));
                        Some(Symbol::Word(source))
                    }
                    _ => None,
                }
            }
            _ => None,
        }
    }
}

#[derive(Debug, PartialEq)]
struct Plan {
    destination: String,
    source: String,
    count: String,
    index: String,
    stride: String,
}

fn truth_variable(expression: &Expression) -> Option<&str> {
    if let Some(name) = ident(expression) {
        return Some(name);
    }
    match &expression.kind {
        ExprKind::Binary {
            op: BinOp::NotEq | BinOp::Integer(IntOp::Or | IntOp::Ne, BitLane::W32),
            left,
            right,
        } if matches!(right.kind, ExprKind::Lit(Literal::Int(0))) => truth_variable(left),
        _ => None,
    }
}

fn match_for(statement: &Statement, context: &Context) -> Option<Plan> {
    let StmtKind::For {
        init,
        cond: Some(cond),
        update: Some(update),
        body,
    } = &statement.kind
    else {
        return None;
    };
    let count = truth_variable(cond)?.to_owned();
    if let Some(init) = init {
        let StmtKind::Expr(Expression {
            kind: ExprKind::Assign { target, .. },
            ..
        }) = &init.kind
        else {
            return None;
        };
        // Lowering retains this entire initializer before the runtime guard;
        // only its destination, not its RHS shape, supplies the count proof.
        if ident(target) != Some(&count) {
            return None;
        }
    }
    let decrements = matches!(&update.kind, ExprKind::Unary { op: UnaryOp::PostDec | UnaryOp::PreDec, expr } if ident(expr) == Some(&count));
    if !decrements || body.len() != 2 || !context.integers.contains(&count) {
        return None;
    }
    let mut evaluation = Evaluation {
        values: context
            .local
            .iter()
            .map(|name| (name.clone(), Symbol::Address(Affine::variable(name))))
            .collect(),
        writes: HashSet::new(),
        reads: vec![],
        stores: vec![],
    };
    for statement in body {
        let StmtKind::Expr(value) = &statement.kind else {
            return None;
        };
        evaluation.evaluate(value)?;
    }
    let [read] = evaluation.reads.as_slice() else {
        return None;
    };
    let [(destination_address, stored)] = evaluation.stores.as_slice() else {
        return None;
    };
    if stored != read {
        return None;
    }
    if read.terms.len() != 1 {
        return None;
    }
    let (source, coefficient) = read.terms.first_key_value()?;
    if *coefficient != 1 {
        return None;
    }
    let source = source.clone();
    if read.offset != 0
        || evaluation.values.get(&source)
            != Some(&Symbol::Address(
                Affine::variable(&source).plus(Affine::constant(2))?,
            ))
    {
        return None;
    }
    if destination_address.offset != 0 || destination_address.terms.len() != 2 {
        return None;
    }
    let destination = destination_address
        .terms
        .iter()
        .find_map(|(name, coefficient)| (*coefficient == 1).then(|| name.clone()))?;
    let index = destination_address
        .terms
        .iter()
        .find_map(|(name, coefficient)| (*coefficient == 2).then(|| name.clone()))?;
    let Symbol::Address(after_index) = evaluation.values.get(&index)? else {
        return None;
    };
    let difference = after_index
        .clone()
        .plus(Affine::variable(&index).scaled(-1)?)?;
    if difference.terms.len() != 1 {
        return None;
    }
    let (stride, coefficient) = difference.terms.first_key_value()?;
    if *coefficient != 1 {
        return None;
    }
    let stride = stride.clone();
    if difference.offset != 0
        || !context.integers.contains(&index)
        || !context.integers.contains(&stride)
    {
        return None;
    }
    let names = [&destination, &source, &count, &index, &stride];
    if names.iter().collect::<HashSet<_>>().len() != names.len()
        || names.iter().any(|name| {
            !context.local.contains(*name)
                || !context.ordinary.contains(*name)
                || context.forbidden.contains(*name)
        })
        || evaluation.writes.contains(&destination)
        || evaluation.writes.contains(&count)
        || evaluation.writes.contains(&stride)
    {
        return None;
    }
    let mut inside = Context::default();
    statement.walk_exprs(&mut |value| inspect(value, &mut inside));
    // Auxiliary writes may disappear only when ordinary, unexposed local
    // storage has no observer outside this loop.
    for name in &evaluation.writes {
        if name != &source
            && name != &index
            && (context.references.get(name) != inside.references.get(name)
                || (context.declared.contains(name)
                    && (!context.local.contains(name)
                        || !context.ordinary.contains(name)
                        || context.forbidden.contains(name))))
        {
            return None;
        }
    }
    Some(Plan {
        destination,
        source,
        count,
        index,
        stride,
    })
}

impl Plan {
    // Evaluated after the unchanged For initializer. Zero iterations never
    // normalize pointers; all other paths require the complete scalar footprint.
    fn guard(&self) -> Expression {
        let wide = |name: &str| Expression::ident(name);
        let add = |left, right| word(IntOp::Add, BitLane::W64, left, right);
        let mul = |left, right| word(IntOp::Mul, BitLane::W64, left, right);
        let last_index = add(
            wide(&self.index),
            mul(
                word(IntOp::Sub, BitLane::W64, wide(&self.count), int(1)),
                wide(&self.stride),
            ),
        );
        let final_index = add(
            wide(&self.index),
            mul(wide(&self.count), wide(&self.stride)),
        );
        let first = add(wide(&self.destination), mul(wide(&self.index), int(2)));
        let last = add(wide(&self.destination), mul(last_index.clone(), int(2)));
        let source_end = add(wide(&self.source), mul(wide(&self.count), int(2)));
        let memory = || mul(call("__c_ptr_memory_size", vec![]), int(65536));
        let numeric_integer = |name: &str| {
            [
                number(name),
                binary(BinOp::GtEq, wide(name), int(i32::MIN as i64)),
                binary(BinOp::LtEq, wide(name), int(i32::MAX as i64)),
                binary(
                    BinOp::Eq,
                    Expression::new(ExprKind::Cast {
                        expr: Box::new(wide(name)),
                        type_name: "double".into(),
                    }),
                    word(IntOp::Or, BitLane::W32, wide(name), int(0)),
                ),
            ]
            .into_iter()
            .reduce(|a, b| binary(BinOp::And, a, b))
            .unwrap()
        };
        let mut bounds = vec![];
        for name in [&self.destination, &self.source, &self.index, &self.stride] {
            bounds.push(numeric_integer(name));
        }
        bounds.push(word(
            IntOp::GeS,
            BitLane::W64,
            wide(&self.destination),
            int(0),
        ));
        for value in [
            last_index.clone(),
            final_index,
            mul(wide(&self.index), int(2)),
            mul(last_index, int(2)),
            mul(wide(&self.stride), int(2)),
        ] {
            bounds.push(word(
                IntOp::GeS,
                BitLane::W64,
                value.clone(),
                int(i32::MIN as i64),
            ));
            bounds.push(word(IntOp::LeS, BitLane::W64, value, int(i32::MAX as i64)));
        }
        for value in [first, last] {
            bounds.push(word(IntOp::GeS, BitLane::W64, value.clone(), int(0)));
            bounds.push(word(
                IntOp::LeS,
                BitLane::W64,
                value.clone(),
                int(i32::MAX as i64),
            ));
            bounds.push(word(IntOp::LeU, BitLane::W64, add(value, int(2)), memory()));
        }
        bounds.push(word(IntOp::GeS, BitLane::W64, wide(&self.source), int(0)));
        bounds.push(word(
            IntOp::LeS,
            BitLane::W64,
            source_end.clone(),
            int(i32::MAX as i64),
        ));
        bounds.push(word(IntOp::LeU, BitLane::W64, source_end, memory()));
        let bounds = bounds
            .into_iter()
            .reduce(|a, b| binary(BinOp::And, a, b))
            .unwrap();
        let access = binary(
            BinOp::Or,
            word(IntOp::Eq, BitLane::W32, wide(&self.count), int(0)),
            bounds,
        );
        binary(
            BinOp::And,
            numeric_integer(&self.count),
            binary(
                BinOp::And,
                word(IntOp::GeS, BitLane::W64, wide(&self.count), int(0)),
                access,
            ),
        )
    }
}

pub(super) fn normalize_for(statement: &mut Statement, context: &Context) -> bool {
    let Some(plan) = match_for(statement, context) else {
        return false;
    };
    let mut scalar = statement.clone();
    let StmtKind::For { init, .. } = &mut scalar.kind else {
        unreachable!();
    };
    let initializer = init.take();
    let n = || Expression::ident(&plan.count);
    let product = |left, right| word(IntOp::Mul, BitLane::W32, left, right);
    let transfer = vec![
        Statement::new(StmtKind::Expr(call(
            "__c_ptr_copy_column16",
            vec![
                word(
                    IntOp::Add,
                    BitLane::W32,
                    Expression::ident(&plan.destination),
                    product(Expression::ident(&plan.index), int(2)),
                ),
                Expression::ident(&plan.source),
                product(Expression::ident(&plan.stride), int(2)),
                n(),
            ],
        ))),
        assign(
            &plan.source,
            word(
                IntOp::Add,
                BitLane::W32,
                Expression::ident(&plan.source),
                product(n(), int(2)),
            ),
        ),
        assign(
            &plan.index,
            word(
                IntOp::Add,
                BitLane::W32,
                Expression::ident(&plan.index),
                product(n(), Expression::ident(&plan.stride)),
            ),
        ),
    ];
    let guarded = Statement::new(StmtKind::If {
        cond: plan.guard(),
        then_body: vec![
            Statement::new(StmtKind::If {
                cond: word(IntOp::Ne, BitLane::W32, n(), int(0)),
                then_body: transfer,
                elifs: vec![],
                else_body: None,
            }),
            assign(&plan.count, int(0)),
        ],
        elifs: vec![],
        // No cursor/index updates move across a potentially trapping scalar load.
        else_body: Some(vec![scalar]),
    });
    let mut body = Vec::with_capacity(2);
    if let Some(initializer) = initializer {
        body.push(*initializer);
    }
    body.push(guarded);
    statement.kind = StmtKind::Block(body);
    true
}

#[test]
fn staged_typed16_for_matcher_keeps_runtime_stride_and_live_bindings() {
    for element in [
        "short",
        "short int",
        "signed short",
        "signed short int",
        "unsigned short",
        "unsigned short int",
    ] {
        let source = format!(
            "void scatter({element} *d, {element} *s, int n, int width) {{ int idx=0; int j; for(j=n;j;j--) {{ d[idx]=*s++; idx+=width; }} }}"
        );
        let expected = format!("{:?}", crate::parse_with_peg(&source).unwrap());
        for backend in [crate::frontend::Backend::Peg, crate::frontend::Backend::Native] {
        let mut module = crate::frontend::with_backend(backend, || crate::parse(&source)).unwrap();
        assert_eq!(format!("{module:?}"), expected, "{backend:?}: {element}");
        assert_eq!(
            word_transfers(&module.body),
            1,
            "typed16 scatter: {element}"
        );
        let before = format!("{:#?}", module.body);
        super::normalize(&mut module.body);
        assert_eq!(
            format!("{:#?}", module.body),
            before,
            "normalization must be idempotent: {element}"
        );
        }
    }
}

#[test]
fn staged_typed16_word_masks_preserve_bits_and_reject_unproved_values() {
    fn evaluate(value: &Expression) -> (Option<Symbol>, Evaluation) {
        let mut evaluation = Evaluation {
            values: [("s".into(), Symbol::Address(Affine::variable("s")))].into(),
            writes: HashSet::new(), reads: vec![], stores: vec![],
        };
        let result = evaluation.evaluate(value);
        (result, evaluation)
    }
    for load in ["__c_ptr_i32_load16_s", "__c_ptr_i32_load16_u"] {
        for mask in [0xffff, 0x1ffff, i64::from(u32::MAX), -1, -65537] {
            for op in [BinOp::BitAnd, BinOp::Integer(IntOp::And, BitLane::W32)] {
                for reversed in [false, true] {
                    let loaded = call(load, vec![Expression::ident("s")]);
                    let masked = if reversed { binary(op, int(mask), loaded) }
                        else { binary(op, loaded, int(mask)) };
                    let (result, evaluation) = evaluate(&masked);
                    assert_eq!(result, Some(Symbol::Word(Affine::variable("s"))));
                    assert_eq!(evaluation.reads, vec![Affine::variable("s")]);
                    assert!(evaluation.writes.is_empty());
                }
            }
            // Independent scalar oracle covers every loaded bit pattern under
            // both signed and unsigned promotion before a 16-bit store.
            for bits in 0..=u16::MAX {
                for promoted in [i64::from(bits), i64::from(bits as i16)] {
                    assert_eq!((promoted & mask) as u16, bits);
                }
            }
        }
    }
    for mask in [0, 0xff, 0x7fff, 0xfffe, i64::MAX] {
        let masked = word(IntOp::And, BitLane::W32,
            call("__c_ptr_i32_load16_u", vec![Expression::ident("s")]), int(mask));
        assert_eq!(evaluate(&masked).0, None, "mask={mask}");
    }
    for value in [
        Expression::ident("s"),
        call("unknown", vec![]),
        call("__c_ptr_i32_load", vec![Expression::ident("s")]),
        call("__c_ptr_i32_load16_u", vec![call("side_effect", vec![])]),
    ] {
        assert_eq!(evaluate(&word(IntOp::And, BitLane::W32, value, int(0xffff))).0, None);
    }
    let loaded = || call("__c_ptr_i32_load16_u", vec![Expression::ident("s")]);
    for value in [
        word(IntOp::And, BitLane::W64, loaded(), int(0xffff)),
        word(IntOp::And, BitLane::W32, loaded(), Expression::ident("mask")),
        word(IntOp::And, BitLane::W32, loaded(), call("mask_once", vec![])),
    ] {
        assert_eq!(evaluate(&value).0, None);
    }
}

#[test]
fn staged_typed16_condition_and_word_cast_match_current_c_lowering() {
    let condition = binary(
        BinOp::NotEq,
        word(IntOp::Or, BitLane::W32, Expression::ident("j"), int(0)),
        int(0),
    );
    assert_eq!(truth_variable(&condition), Some("j"));
    assert_eq!(
        truth_variable(&binary(BinOp::Eq, Expression::ident("j"), int(0))),
        None
    );
    let mut evaluation = Evaluation {
        values: [("s".into(), Symbol::Address(Affine::variable("s")))].into(),
        writes: HashSet::new(),
        reads: vec![],
        stores: vec![],
    };
    let loaded = Expression::new(ExprKind::Cast {
        type_name: "short".into(),
        expr: Box::new(call("__c_ptr_i32_load16_s", vec![Expression::ident("s")])),
    });
    assert_eq!(
        evaluation.evaluate(&loaded),
        Some(Symbol::Word(Affine::variable("s")))
    );
}

#[test]
fn staged_typed16_fixture_local_cursors_and_surrounding_loops_select_transfer() {
    let module = crate::parse(include_str!(
        "../../../tests/c/doom_startup/typed16_for_column.c"
    ))
    .expect("complete multi-loop fixture AST");
    assert_eq!(
        word_transfers(&module.body),
        1,
        "local short cursors in the complete fixture must select one transfer"
    );
}

#[cfg(test)]
fn word_transfers(body: &[Statement]) -> usize {
    let mut count = 0;
    for statement in body {
        statement.walk_exprs(&mut |value| {
            count += usize::from(matches!(&value.kind, ExprKind::Call { callee, .. } if ident(callee) == Some("__c_ptr_copy_column16")));
        });
    }
    count
}

#[test]
fn staged_typed16_lowering_keeps_initializer_once_and_scalar_postfix_order() {
    let mut context = Context::default();
    for name in ["d", "s", "j", "idx", "width", "n"] {
        context.local.insert(name.into());
        context.declared.insert(name.into());
        context.ordinary.insert(name.into());
    }
    for name in ["j", "idx", "width", "n"] {
        context.integers.insert(name.into());
    }
    let source = Expression::new(ExprKind::Sequence(vec![
        Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident("saved")),
            value: Box::new(Expression::ident("s")),
        }),
        Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident("s")),
            value: Box::new(word(
                IntOp::Add,
                BitLane::W32,
                Expression::ident("s"),
                int(2),
            )),
        }),
        call("__c_ptr_i32_load16_s", vec![Expression::ident("saved")]),
    ]));
    let mut original = Statement::new(StmtKind::For {
        init: Some(Box::new(assign("j", Expression::ident("n")))),
        cond: Some(Expression::ident("j")),
        update: Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::PostDec,
            expr: Box::new(Expression::ident("j")),
        })),
        body: vec![
            Statement::new(StmtKind::Expr(call(
                "__c_ptr_i32_store16",
                vec![
                    word(
                        IntOp::Add,
                        BitLane::W32,
                        Expression::ident("d"),
                        word(IntOp::Mul, BitLane::W32, Expression::ident("idx"), int(2)),
                    ),
                    source,
                ],
            ))),
            assign(
                "idx",
                word(
                    IntOp::Add,
                    BitLane::W32,
                    Expression::ident("idx"),
                    Expression::ident("width"),
                ),
            ),
        ],
    });
    original.walk_exprs(&mut |value| inspect(value, &mut context));
    let mut lowered = original.clone();
    assert!(normalize_for(&mut lowered, &context));
    let StmtKind::For { init, .. } = &mut original.kind else {
        unreachable!();
    };
    let initializer = init.take().unwrap();
    let StmtKind::Block(body) = &lowered.kind else {
        panic!("outer block");
    };
    assert_eq!(body.len(), 2);
    assert_eq!(format!("{:?}", body[0]), format!("{:?}", initializer));
    let StmtKind::If {
        else_body: Some(fallback),
        ..
    } = &body[1].kind
    else {
        panic!("guard");
    };
    assert_eq!(fallback.len(), 1);
    assert_eq!(format!("{:?}", fallback[0]), format!("{:?}", original));
    assert_eq!(word_transfers(body), 1);
}

#[test]
fn staged_typed16_for_rejects_effects_volatile_and_exposed_state() {
    for extra in ["int *alias=&idx;", "int *alias=&width;", "int *alias=&j;"] {
        let source = format!(
            "void f(short *d, short *s, int n, int width) {{ int idx=0; int j; {extra} for(j=n;j;j--) {{ d[idx]=*s++; idx+=width; }} }}"
        );
        let module = crate::parse(&source).unwrap();
        assert_eq!(word_transfers(&module.body), 0, "{source}");
    }
    for source in [
        "void f(volatile short *d, short *s, int n, int width) { int idx=0; int j; for(j=n;j;j--) { d[idx]=*s++; idx+=width; } }",
        "void f(short *d, short *s, int n, int width) { int idx=0; int j; for(j=n;j;j--) { d[idx]=*s++; idx+=width++; } }",
        "void f(short *d, short *s, int n, int width) { int idx=0; int j; for(j=n;j;j--) { d[idx]=*s++ + 1; idx+=width; } }",
        "void f(unsigned char *d, unsigned char *s, int n, int width) { int idx=0; int j; for(j=n;j;j--) { d[idx]=*s++; idx+=width; } }",
        "void f(unsigned short *d, unsigned short *s, int n, int width) { int idx=0; int j; for(j=n;j;j--) { d[idx]=*s++ & 0x7fff; idx+=width; } }",
        "void f(unsigned short *d, unsigned short *s, int n, int width, int mask) { int idx=0; int j; for(j=n;j;j--) { d[idx]=*s++ & mask; idx+=width; } }",
    ] {
        for backend in [crate::frontend::Backend::Peg, crate::frontend::Backend::Native] {
            let module = crate::frontend::with_backend(backend, || crate::parse(source)).unwrap();
            assert_eq!(word_transfers(&module.body), 0, "{backend:?}: {source}");
        }
    }
}

// Independent guard interpreter: reject overflow rather than hiding it with
// wrapping arithmetic, and record reads to verify zero-count short circuit.
#[cfg(test)]
fn guard_value(
    expression: &Expression,
    values: &HashMap<String, i64>,
    reads: &mut Vec<String>,
) -> i64 {
    match &expression.kind {
        ExprKind::Ident(name) => {
            reads.push(name.clone());
            values[name]
        }
        ExprKind::Lit(Literal::Int(value)) => *value,
        ExprKind::IsType { .. } => 1,
        ExprKind::Cast { expr, type_name } if type_name == "double" => {
            guard_value(expr, values, reads)
        }
        ExprKind::Call { callee, args, .. }
            if ident(callee) == Some("__c_ptr_memory_size") && args.is_empty() =>
        {
            1
        }
        ExprKind::Binary { op, left, right } => {
            let a = guard_value(left, values, reads);
            if *op == BinOp::And && a == 0 {
                return 0;
            }
            if *op == BinOp::Or && a != 0 {
                return 1;
            }
            let b = guard_value(right, values, reads);
            match op {
                BinOp::And | BinOp::Or => i64::from(b != 0),
                BinOp::Eq | BinOp::Integer(IntOp::Eq, _) => i64::from(a == b),
                BinOp::GtEq | BinOp::Integer(IntOp::GeS, _) => i64::from(a >= b),
                BinOp::LtEq | BinOp::Integer(IntOp::LeS, _) => i64::from(a <= b),
                BinOp::Integer(IntOp::LeU, _) => i64::from((a as u64) <= (b as u64)),
                BinOp::Integer(IntOp::Add, BitLane::W64) => a.checked_add(b).unwrap(),
                BinOp::Integer(IntOp::Sub, BitLane::W64) => a.checked_sub(b).unwrap(),
                BinOp::Integer(IntOp::Mul, BitLane::W64) => a.checked_mul(b).unwrap(),
                BinOp::Integer(IntOp::Or, BitLane::W32) => ((a as i32) | (b as i32)) as i64,
                _ => panic!("unexpected guard operation {op:?}"),
            }
        }
        _ => panic!("unexpected guard expression"),
    }
}

#[test]
fn staged_typed16_guard_checks_complete_words_and_widened_final_state() {
    let plan = Plan {
        destination: "d".into(),
        source: "s".into(),
        count: "j".into(),
        index: "idx".into(),
        stride: "width".into(),
    };
    for (destination, source, count, index, stride, expected) in [
        (100, 200, 8, 0, 3, true),
        (100, 200, 8, 0, -3, true),
        (100, 200, 8, 0, 0, true),
        (65535, 200, 1, 0, 1, false),
        (100, 65535, 1, 0, 1, false),
        (100, 200, 8, 0, -20, false),
        (100, 200, -1, 0, 1, false),
        (100, 200, i32::MAX as i64, 0, i32::MAX as i64, false),
        (100, 200, 1, 0, i32::MAX as i64, false),
    ] {
        let values = [
            ("d", destination),
            ("s", source),
            ("j", count),
            ("idx", index),
            ("width", stride),
        ]
        .into_iter()
        .map(|(k, v)| (k.into(), v))
        .collect();
        assert_eq!(
            guard_value(&plan.guard(), &values, &mut vec![]) != 0,
            expected
        );
    }
}

#[test]
fn staged_typed16_zero_count_does_not_read_or_normalize_pointer_state() {
    let plan = Plan {
        destination: "d".into(),
        source: "s".into(),
        count: "j".into(),
        index: "idx".into(),
        stride: "width".into(),
    };
    let values = [("j".into(), 0)].into();
    let mut reads = vec![];
    assert_eq!(guard_value(&plan.guard(), &values, &mut reads), 1);
    assert!(reads.iter().all(|name| name == "j"));
}
