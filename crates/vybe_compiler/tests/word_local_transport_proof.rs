use vybe_ast::{
    BinOp, BindingPattern, BitLane, ExprKind, Expression, IntOp, LambdaBody, Literal, Statement,
    StmtKind, StorageAccess, TypeBinding, TypeHint, UnaryOp, VarDeclKind, VarDeclarator,
};
use vybe_compiler::word_local_proof::{WordInterpretation, prove as prove_region};

fn prove(
    enabled: bool,
    body: &[Statement],
    resolve: impl Fn(&TypeHint) -> Option<WordInterpretation>,
) -> Option<Vec<vybe_compiler::word_local_proof::BindingProof>> {
    prove_region(
        &vybe_ast::Directives {
            word_local_transport: enabled.then_some(vybe_ast::WordLocalTransport::ProvenW32),
            ..Default::default()
        },
        body,
        resolve,
    )
}

fn lit(value: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Int(value)))
}

fn ident(name: &str) -> Expression {
    Expression::new(ExprKind::Ident(name.into()))
}

fn binary(op: IntOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Integer(op, BitLane::W32),
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn declaration(name: &str, init: Option<Expression>) -> Statement {
    let mut hint = TypeHint::checked("word");
    hint.access = Some(StorageAccess::Ordinary);
    hint.word_value = Some(WordInterpretation::Signed);
    Statement::new(StmtKind::VarDecl {
        kind: VarDeclKind::Var,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.into()),
            type_hint: Some(hint),
            init,
            array_bounds: None,
            with_events: false,
        }],
    })
}

fn compile_function(body: Vec<Statement>, enabled: bool) -> Vec<vybe_runtime::Chunk> {
    let module = vybe_ast::Module {
        name: "boxed-word-proof".into(),
        language: vybe_ast::Lang::Unknown,
        imports: vec![],
        canon: Default::default(),
        directives: vybe_ast::Directives {
            word_local_transport: Some(if enabled {
                vybe_ast::WordLocalTransport::ProvenW32
            } else {
                vybe_ast::WordLocalTransport::Boxed
            }),
            ..Default::default()
        },
        body: vec![
            Statement::new(StmtKind::FunctionDecl {
                name: "word_loop".into(),
                params: vec![],
                return_type: None,
                body,
                modifiers: Default::default(),
                handles: vec![],
                is_async: false,
                is_generator: false,
                is_sub: false,
            }),
            Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Call {
                callee: Box::new(ident("word_loop")),
                args: vec![],
                optional: false,
            })))),
        ],
    };
    vybe_compiler::Compiler::with_profile(
        vybe_compiler::profile::parse_profile("[compiler]\n").unwrap(),
    )
    .compile(&module)
    .unwrap()
}

fn conversions(chunks: &[vybe_runtime::Chunk]) -> usize {
    vybe_runtime::debug::disassemble(
        chunks
            .iter()
            .find(|chunk| chunk.name == "word_loop")
            .unwrap(),
    )
    .matches("i64.trunc_f64_s")
    .count()
}

#[test]
fn certified_boxed_binding_reuses_normalization_without_raw_storage() {
    let baseline = compile_function(region(), false);
    let optimized = compile_function(region(), true);
    assert!(conversions(&optimized) < conversions(&baseline));
    for chunks in [baseline, optimized] {
        assert!(
            chunks
                .iter()
                .all(|chunk| !vybe_runtime::debug::disassemble(chunk).contains("dyneq"))
        );
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            &chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: false,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        assert_eq!(vm.run(chunks).unwrap().as_i32(), 48);
    }
}

#[test]
fn uncertified_and_unsigned_bindings_keep_existing_narrowing() {
    for interpretation in [None, Some(WordInterpretation::Unsigned)] {
        let mut body = region();
        for statement in &mut body {
            if let StmtKind::VarDecl { declarations, .. } = &mut statement.kind {
                for declaration in declarations {
                    declaration.type_hint.as_mut().unwrap().word_value = interpretation;
                }
            }
        }
        assert_eq!(
            conversions(&compile_function(body.clone(), false)),
            conversions(&compile_function(body, true))
        );
    }
}

#[test]
fn changing_declared_shape_invalidates_word_value_certification() {
    let mut hint = TypeHint::checked("arbitrary-certified-spelling");
    hint.word_value = Some(WordInterpretation::Signed);
    hint.push_str("[]");
    assert_eq!(hint.word_value, None);
    hint.word_value = Some(WordInterpretation::Signed);
    hint.set_spelling("other-shape");
    assert_eq!(hint.word_value, None);
}

fn assign(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
        target: Box::new(ident(name)),
        value: Box::new(value),
    })))
}

fn resolve(hint: &TypeHint) -> Option<WordInterpretation> {
    match hint.spelling() {
        "word" => Some(WordInterpretation::Signed),
        "uword" => Some(WordInterpretation::Unsigned),
        _ => None,
    }
}

fn region() -> Vec<Statement> {
    vec![
        declaration("acc", Some(lit(0))),
        declaration("remaining", None),
        assign("remaining", lit(16)),
        Statement::new(StmtKind::While {
            cond: binary(IntOp::GtS, ident("remaining"), lit(0)),
            body: vec![
                assign("acc", binary(IntOp::Add, ident("acc"), lit(3))),
                assign("remaining", binary(IntOp::Sub, ident("remaining"), lit(1))),
            ],
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(ident("acc")))),
    ]
}

#[test]
fn closed_word_loop_requires_explicit_opt_in_and_proves_all_bindings() {
    let body = region();
    assert!(prove(false, &body, resolve).is_none());
    let bindings = prove(true, &body, resolve).unwrap();
    assert_eq!(bindings.len(), 2);
    assert_eq!(bindings[0].name, "acc");
    assert_eq!(bindings[1].name, "remaining");
    assert!(
        bindings
            .iter()
            .all(|binding| binding.interpretation == WordInterpretation::Signed)
    );
}

#[test]
fn compile_result_does_not_issue_physical_storage_proofs_by_default() {
    assert!(
        vybe_compiler::primitives::CompileResult::default()
            .word_local_proofs
            .is_empty()
    );
}

#[test]
fn definite_initialization_intersects_branches_and_rejects_zero_iteration_assumptions() {
    let conditional = |otherwise| {
        Statement::new(StmtKind::If {
            cond: binary(IntOp::Eq, lit(1), lit(1)),
            then_body: vec![assign("x", lit(4))],
            elifs: vec![],
            else_body: otherwise,
        })
    };
    let mut body = vec![
        declaration("x", None),
        conditional(Some(vec![assign("x", lit(7))])),
        Statement::new(StmtKind::Return(Some(ident("x")))),
    ];
    assert!(prove(true, &body, resolve).is_some());
    body[1] = conditional(None);
    assert!(prove(true, &body, resolve).is_none());
    body[1] = Statement::new(StmtKind::While {
        cond: lit(0),
        body: vec![assign("x", lit(4))],
        else_body: None,
    });
    assert!(prove(true, &body, resolve).is_none());
    body[1] = conditional(Some(vec![assign(
        "x",
        Expression::new(ExprKind::Lit(Literal::Float(1.5))),
    )]));
    assert!(prove(true, &body, resolve).is_none());
}

#[test]
fn nonword_write_address_capture_call_and_pointer_carrier_reject_region() {
    let forbidden = vec![
        Expression::new(ExprKind::Lit(Literal::Float(3.5))),
        Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(ident("acc")),
        }),
        Expression::new(ExprKind::Unary {
            op: UnaryOp::Deref,
            expr: Box::new(ident("acc")),
        }),
        Expression::new(ExprKind::Call {
            callee: Box::new(ident("callback")),
            args: vec![],
            optional: false,
        }),
        Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Expr(Box::new(ident("acc"))),
            is_async: false,
            captures: vec!["acc".into()],
        }),
        Expression::new(ExprKind::Binary {
            op: BinOp::Integer(IntOp::Add, BitLane::W64),
            left: Box::new(lit(1)),
            right: Box::new(lit(2)),
        }),
    ];
    for value in forbidden {
        let mut body = region();
        body.insert(2, assign("acc", value));
        assert!(prove(true, &body, resolve).is_none());
    }
}

#[test]
fn unknown_volatile_atomic_descriptive_converting_and_pointer_declarations_reject() {
    for access in [
        None,
        Some(StorageAccess::Unknown),
        Some(StorageAccess::Volatile),
        Some(StorageAccess::Atomic),
    ] {
        let mut body = region();
        let StmtKind::VarDecl { declarations, .. } = &mut body[0].kind else {
            unreachable!()
        };
        declarations[0].type_hint.as_mut().unwrap().access = access;
        assert!(prove(true, &body, resolve).is_none());
    }
    for binding in [TypeBinding::Descriptive, TypeBinding::Converting] {
        let mut body = region();
        let StmtKind::VarDecl { declarations, .. } = &mut body[0].kind else {
            unreachable!()
        };
        declarations[0].type_hint.as_mut().unwrap().binding = binding;
        assert!(prove(true, &body, resolve).is_none());
    }
    let mut body = region();
    let StmtKind::VarDecl { declarations, .. } = &mut body[0].kind else {
        unreachable!()
    };
    declarations[0]
        .type_hint
        .as_mut()
        .unwrap()
        .set_spelling("word*");
    assert!(prove(true, &body, resolve).is_none());
}

#[test]
fn unsigned_interpretation_is_explicit_not_inferred_from_variable_name() {
    let mut body = vec![
        declaration("ordinary", Some(lit(i64::from(u32::MAX)))),
        Statement::new(StmtKind::Return(Some(binary(
            IntOp::ShrU,
            ident("ordinary"),
            lit(1),
        )))),
    ];
    let StmtKind::VarDecl { declarations, .. } = &mut body[0].kind else {
        unreachable!()
    };
    declarations[0]
        .type_hint
        .as_mut()
        .unwrap()
        .set_spelling("uword");
    assert_eq!(
        prove(true, &body, resolve).unwrap()[0].interpretation,
        WordInterpretation::Unsigned
    );
    assert!(prove(true, &body, |_| None).is_none());
    body.insert(
        1,
        assign("ordinary", binary(IntOp::Add, ident("ordinary"), lit(1))),
    );
    // A signed W32 result cannot silently become an unsigned value carrier.
    assert!(prove(true, &body, resolve).is_none());
}

#[test]
fn storage_interpretation_and_declaration_order_are_part_of_the_proof() {
    let body = vec![declaration("x", Some(lit(i64::from(u32::MAX))))];
    assert!(prove(true, &body, resolve).is_none());
    let body = vec![assign("x", lit(1)), declaration("x", None)];
    assert!(prove(true, &body, resolve).is_none());
    let body = vec![declaration("x", Some(ident("x")))];
    assert!(prove(true, &body, resolve).is_none());
    let body = vec![declaration("x", Some(binary(IntOp::Eq, lit(1), lit(1))))];
    // W32 comparison can have a boolean value ABI; bit width is not enough.
    assert!(prove(true, &body, resolve).is_none());
}

#[test]
fn duplicate_shadowed_and_static_bindings_are_not_physical_slot_proofs() {
    let mut body = region();
    body.insert(1, declaration("acc", Some(lit(0))));
    assert!(prove(true, &body, resolve).is_none());
    let mut body = region();
    body.insert(
        2,
        Statement::new(StmtKind::Block(vec![declaration("acc", Some(lit(0)))])),
    );
    assert!(prove(true, &body, resolve).is_none());
    let mut body = region();
    let StmtKind::VarDecl { kind, .. } = &mut body[0].kind else {
        unreachable!()
    };
    *kind = VarDeclKind::Static;
    assert!(prove(true, &body, resolve).is_none());
}

fn subset(body: &[Statement]) -> vybe_compiler::word_local_proof::SubsetProof {
    vybe_compiler::word_local_proof::prove_subset(
        &vybe_ast::Directives {
            word_local_transport: Some(vybe_ast::WordLocalTransport::ProvenW32),
            ..Default::default()
        },
        body,
        resolve,
        |name| resolve(&TypeHint::checked(name)),
    )
}

#[test]
fn subset_preserves_unrelated_pointer_locals_fields_and_pixel_stores() {
    let mut body = region();
    body.insert(
        0,
        Statement::new(StmtKind::VarDecl {
            kind: VarDeclKind::Let,
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("buffer".into()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
        }),
    );
    body.insert(
        4,
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(ident("buffer")),
                index: Box::new(ident("acc")),
                null_safe: false,
            })),
            value: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(ident("palette")),
                field: "pixel".into(),
                null_safe: false,
            })),
        }))),
    );
    assert!(prove(true, &body, resolve).is_none());
    let proof = subset(&body);
    assert!(proof.rejected.is_empty());
    assert_eq!(proof.eligible.len(), 2);
    assert!(
        !proof
            .eligible
            .iter()
            .any(|binding| binding.name == "buffer")
    );
}

#[test]
fn subset_rejects_only_escaping_candidate_and_keeps_independent_word_local() {
    let mut body = region();
    body.insert(
        4,
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(ident("acc")),
        }))),
    );
    let proof = subset(&body);
    assert_eq!(proof.eligible.len(), 1);
    assert_eq!(proof.eligible[0].name, "remaining");
    assert_eq!(
        proof.rejected[0].reason,
        vybe_compiler::word_local_proof::Rejection::AddressOrReferenceEscape
    );
}

#[test]
fn subset_unsigned_store_requires_certified_finite_word_cast_not_helper_name() {
    let mut declaration = declaration("bits", Some(lit(0)));
    let StmtKind::VarDecl { declarations, .. } = &mut declaration.kind else {
        unreachable!()
    };
    declarations[0]
        .type_hint
        .as_mut()
        .unwrap()
        .set_spelling("uword");
    let raw = binary(IntOp::Add, ident("bits"), lit(1));
    let mut body = vec![
        declaration,
        assign(
            "bits",
            Expression::new(ExprKind::Cast {
                expr: Box::new(raw.clone()),
                type_name: "uword".into(),
            }),
        ),
    ];
    assert_eq!(
        subset(&body).eligible[0].interpretation,
        WordInterpretation::Unsigned
    );
    body[1] = assign("bits", raw);
    assert!(subset(&body).eligible.is_empty());
    body[1] = assign(
        "bits",
        Expression::new(ExprKind::Cast {
            expr: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(ident("number")),
                args: vec![],
                optional: false,
            })),
            type_name: "uword".into(),
        }),
    );
    assert!(subset(&body).eligible.is_empty());
}

#[test]
fn subset_incompatible_join_and_capture_do_not_certify_storage() {
    let mut body = region();
    body.insert(
        4,
        Statement::new(StmtKind::If {
            cond: ident("external_condition"),
            then_body: vec![assign("acc", binary(IntOp::Add, ident("acc"), lit(1)))],
            elifs: vec![],
            else_body: Some(vec![assign(
                "acc",
                Expression::new(ExprKind::Lit(Literal::Float(1.5))),
            )]),
        }),
    );
    assert_eq!(
        subset(&body).rejected[0].reason,
        vybe_compiler::word_local_proof::Rejection::IncompatibleProducer
    );
    body[4] = Statement::new(StmtKind::Expr(Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Expr(Box::new(lit(0))),
        is_async: false,
        captures: vec!["acc".into()],
    })));
    assert_eq!(
        subset(&body).rejected[0].reason,
        vybe_compiler::word_local_proof::Rejection::Capture
    );
}
