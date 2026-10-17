use super::super::{AtomicModuleScan, collect_atomic_place_idents, stmt_uses_proxy};
use super::*;
use vybe_ast::*;

fn stmt(kind: StmtKind) -> Statement {
    Statement::new(kind)
}
fn function(name: &str, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::FunctionDecl {
        name: name.into(),
        params: Vec::new(),
        return_type: Some("number".into()),
        body,
        modifiers: Default::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}
fn address(name: &str) -> Statement {
    stmt(StmtKind::Expr(Expression::new(ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr: Box::new(Expression::ident(name)),
    })))
}
fn load(name: &str) -> Statement {
    stmt(StmtKind::Expr(Expression::new(ExprKind::Atomic(
        AtomicOp::Load {
            place: Box::new(Expression::ident(name)),
            ordering: MemoryOrder::SeqCst,
        },
    ))))
}
fn features() -> SourceFeaturePreparation {
    SourceFeaturePreparation::new(false, true, false)
}

#[test]
fn typed_unit_transaction_preserves_statement_and_expression_storage_through_seal() {
    let body = vec![
        address("global"),
        function("later", vec![stmt(StmtKind::Empty)]),
        stmt(StmtKind::Expr(Expression::int(7))),
    ];
    let address = body.as_ptr();
    let capacity = body.capacity();
    let statements = body
        .iter()
        .map(|statement| statement as *const Statement)
        .collect::<Vec<_>>();
    let expressions = body
        .iter()
        .map(|statement| match &statement.kind {
            StmtKind::Expr(value) => Some(value as *const Expression),
            _ => None,
        })
        .collect::<Vec<_>>();
    let mut preparation = TypedUnitPreparation::new(
        Cow::Owned(body),
        None,
        HashSet::new(),
        HashSet::new(),
        features(),
    );
    let mut collected = Vec::new();
    while let Some(index) = preparation.collect_next() {
        collected.push(index);
        assert_eq!(
            &preparation.body[index] as *const Statement,
            statements[index]
        );
    }
    assert_eq!(collected, [0, 1, 2]);
    assert!(preparation.collect_next().is_none());
    let prepared = preparation.seal().unwrap();
    let Cow::Owned(body) = prepared.body else {
        panic!("owned source became borrowed")
    };
    assert_eq!(body.as_ptr(), address);
    assert_eq!(body.capacity(), capacity);
    for (index, statement) in body.iter().enumerate() {
        assert_eq!(statement as *const Statement, statements[index]);
        if let StmtKind::Expr(value) = &statement.kind {
            assert_eq!(Some(value as *const Expression), expressions[index]);
        }
    }
    assert_eq!(prepared.references, HashSet::from(["global".into()]));
    assert_eq!(prepared.functions[0].name, "later");
}

#[test]
fn typed_unit_transaction_matches_separate_whole_source_fact_discovery() {
    let body = vec![
        function("early", vec![address("Global")]),
        stmt(StmtKind::Expr(Expression::new(ExprKind::Lit(
            Literal::BigInt(3),
        )))),
        stmt(StmtKind::Expr(Expression::new(ExprKind::Proxy {
            target: Box::new(Expression::ident("target")),
            handler: Box::new(Expression::ident("handler")),
        }))),
        function("late", vec![load("atomic_place")]),
    ];
    for fold in [None, Some(CaseAlphabet::Ascii), Some(CaseAlphabet::Unicode)] {
        for bigint_seed in [false, true] {
            for inspect_proxy in [false, true] {
                for proxy_seed in [false, true] {
                    let atomic_seed = HashSet::from(["previous_atomic".into()]);
                    let receipt_seed = HashSet::from([(13, 2)]);
                    let references =
                        super::super::reference_analysis::module_address_taken(&body, fold);
                    let mut atomic_names = atomic_seed.clone();
                    let mut source_bodies = receipt_seed.clone();
                    collect_atomic_place_idents(
                        &body,
                        &mut AtomicModuleScan {
                            class_depth: 0,
                            names: &mut atomic_names,
                            bodies: &mut source_bodies,
                        },
                    );
                    if !atomic_names.is_empty() {
                        source_bodies.clear();
                    }
                    let bigint_enabled = bigint_seed
                        || body.iter().any(|statement| {
                            let mut found = false;
                            statement.walk_exprs(&mut |value| {
                                found |= matches!(value.kind, ExprKind::Lit(Literal::BigInt(_)));
                            });
                            found
                        });
                    let uses_proxy =
                        proxy_seed || (inspect_proxy && body.iter().any(stmt_uses_proxy));
                    let mut functions = FunctionRegistrationPreparation::default();
                    functions.consume(&body);
                    let functions = format!("{:?}", functions.finish());
                    let mut transaction = TypedUnitPreparation::new(
                        Cow::Borrowed(&body),
                        fold,
                        atomic_seed,
                        receipt_seed,
                        SourceFeaturePreparation::new(bigint_seed, inspect_proxy, proxy_seed),
                    );
                    while transaction.collect_next().is_some() {}
                    let prepared = transaction.seal().unwrap();
                    assert_eq!(prepared.references, references);
                    assert_eq!(prepared.atomic_names, atomic_names);
                    assert_eq!(prepared.source_bodies, source_bodies);
                    assert_eq!(
                        prepared.features,
                        SourceFeatures {
                            bigint_enabled,
                            uses_proxy
                        }
                    );
                    assert_eq!(format!("{:?}", prepared.functions), functions);
                    assert!(matches!(prepared.body, Cow::Borrowed(_)));
                    assert_eq!(prepared.body.as_ptr(), body.as_ptr());
                }
            }
        }
    }
}

#[test]
fn typed_unit_transaction_retains_only_valid_stable_source_receipts() {
    for late_atomic in [false, true] {
        let mut body = vec![function("first", vec![stmt(StmtKind::Empty)])];
        if late_atomic {
            body.push(load("late_place"));
        }
        let StmtKind::FunctionDecl {
            body: procedure, ..
        } = &body[0].kind
        else {
            unreachable!()
        };
        let identity = (procedure.as_ptr() as usize, procedure.len());
        let mut transaction = TypedUnitPreparation::new(
            Cow::Owned(body),
            None,
            HashSet::new(),
            HashSet::new(),
            features(),
        );
        while transaction.collect_next().is_some() {}
        let prepared = transaction.seal().unwrap();
        if late_atomic {
            assert!(prepared.source_bodies.is_empty());
        } else {
            assert!(prepared.source_bodies.contains(&identity));
        }
    }
}

#[test]
fn typed_unit_transaction_cannot_seal_partial_or_abandoned_source() {
    for collected in [0, 1] {
        let body = vec![
            address("unpublished"),
            function("also_unpublished", Vec::new()),
        ];
        let mut transaction = TypedUnitPreparation::new(
            Cow::Owned(body),
            None,
            HashSet::new(),
            HashSet::new(),
            features(),
        );
        for _ in 0..collected {
            transaction.collect_next().unwrap();
        }
        assert!(transaction.seal().err().unwrap().contains("incomplete"));
    }
    let mut abandoned = TypedUnitPreparation::new(
        Cow::Owned(vec![address("failed_source")]),
        None,
        HashSet::new(),
        HashSet::new(),
        features(),
    );
    abandoned.collect_next();
    drop(abandoned);
    let empty = TypedUnitPreparation::new(
        Cow::Owned(Vec::new()),
        None,
        HashSet::new(),
        HashSet::new(),
        features(),
    )
    .seal()
    .unwrap();
    assert!(empty.references.is_empty() && empty.functions.is_empty() && empty.body.is_empty());
}

#[test]
fn typed_unit_function_facts_survive_partial_merge_without_reconstruction() {
    let partial = |members| {
        stmt(StmtKind::ClassDecl {
            name: "Partial".into(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers {
                is_partial: true,
                ..Default::default()
            },
            decorators: Vec::new(),
        })
    };
    let body = vec![
        function("before", vec![address("global")]),
        partial(Vec::new()),
        function("after", Vec::new()),
        partial(vec![ClassMember::Method(Box::new(function(
            "method",
            Vec::new(),
        )))]),
    ];
    let mut transaction = TypedUnitPreparation::new(
        Cow::Owned(body),
        None,
        HashSet::new(),
        HashSet::new(),
        features(),
    );
    while transaction.collect_next().is_some() {}
    let prepared = transaction.seal().unwrap();
    let merged = super::super::merge_partial_classes(&prepared.body, true);
    let mut expected = FunctionRegistrationPreparation::default();
    expected.consume(&merged);
    assert_eq!(
        format!("{:?}", prepared.functions),
        format!("{:?}", expected.finish())
    );
    assert_eq!(
        prepared
            .functions
            .iter()
            .map(|function| function.name.as_str())
            .collect::<Vec<_>>(),
        ["before", "after"]
    );
}
