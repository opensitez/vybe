use vybe_ast::*;
use vybe_compiler::primitives::{CompileResult, Compiler};

fn profile() -> vybe_runtime::profile::LanguageProfile {
    let mut profile = vybe_compiler::profile::parse_profile(
        "[info]\nname='prepared-unit-test'\n[compiler]\n",
    ).unwrap();
    profile.ecma_strict_mode = true;
    profile
}

fn int(value: i64) -> Expression { Expression::new(ExprKind::Lit(Literal::Int(value))) }
fn name(value: &str) -> Expression { Expression::new(ExprKind::Ident(value.into())) }
fn call(value: &str) -> Expression {
    Expression::new(ExprKind::Call { callee: Box::new(name(value)), args: Vec::new(), optional: false })
}
fn function(name: &str, params: Vec<Param>, body: Vec<Statement>) -> Statement {
    Statement::new(StmtKind::FunctionDecl { name: name.into(), params, return_type: None,
        body, modifiers: Default::default(), handles: Vec::new(),
        is_async: false, is_generator: false, is_sub: false })
}
fn fixture() -> Module {
    let default = Param { name: "value".into(), type_hint: None, default: Some(int(7)),
        pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_optional: true, is_nullable: false };
    let mut body = vec![
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Lit(Literal::Str("use strict".into()))))),
        function("first", vec![default], vec![Statement::new(StmtKind::Return(Some(name("value"))))]),
        Statement::new(StmtKind::VarDecl { declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("answer".into()), type_hint: None,
            init: Some(call("later")), array_bounds: None, with_events: false,
        }], kind: VarDeclKind::Let }),
        function("later", Vec::new(), vec![Statement::new(StmtKind::Return(Some(call("first"))))]),
        Statement::new(StmtKind::Empty),
        Statement::new(StmtKind::Return(Some(name("answer")))),
    ];
    for (index, statement) in body.iter_mut().enumerate() {
        statement.span = Span { start_line: index as u32 + 1, start_col: 1,
            end_line: index as u32 + 1, end_col: 20 };
    }
    Module { name: "prepared-source".into(), language: Lang::Unknown, body,
        imports: Vec::new(), directives: Default::default(), canon: Default::default() }
}

fn assert_same(expected: &CompileResult, actual: &CompileResult) {
    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
    let options = vybe_platform_wasm::WasmWriteOptions {
        include_vybe_metadata: false, include_name_section: true,
    };
    let expected = vybe_platform_wasm::write_wasm_with_options(&expected.chunks, options);
    let actual = vybe_platform_wasm::write_wasm_with_options(&actual.chunks,
        vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
    assert_eq!(actual, expected);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&actual).unwrap();
}

#[test]
fn owned_units_preserve_defaults_forward_calls_lines_and_all_compile_metadata() {
    let source = fixture();
    let original = format!("{source:?}");
    let expected = Compiler::with_profile(profile()).compile_with_imports(&source).unwrap();
    let mut automatic = Compiler::with_profile(profile()).begin_emission_owned(source.clone()).unwrap();
    let mut automatic_order = Vec::new();
    while let Some(progress) = automatic.emit_next().unwrap() {
        automatic_order.push((progress.source_index, progress.function_declaration));
    }
    assert_eq!(automatic_order, [(1, true), (3, true), (0, false), (2, false), (4, false), (5, false)]);
    assert_same(&expected, &automatic.finish().unwrap());

    let owned = source.clone();
    let statement_addresses = owned.body.iter().map(|statement| statement as *const Statement as usize).collect::<Vec<_>>();
    let expression_addresses = owned.body.iter().map(|statement| match &statement.kind {
        StmtKind::Expr(value) => Some(value as *const Expression as usize), _ => None,
    }).collect::<Vec<_>>();
    let mut session = Compiler::with_profile(profile()).begin_emission_owned(owned).unwrap();
    let mut producer = session.take_owned_units().unwrap();
    assert_eq!(producer.retained_statements(), source.body.len());
    let mut order = Vec::new();
    while let Some(unit) = producer.next_unit() {
        let index = unit.source_index();
        assert_eq!(unit.statement() as *const Statement as usize, statement_addresses[index]);
        if let StmtKind::Expr(value) = &unit.statement().kind {
            assert_eq!(Some(value as *const Expression as usize), expression_addresses[index]);
        }
        let function = unit.is_function_declaration();
        let progress = session.emit_unit(unit).unwrap();
        assert_eq!((progress.source_index, progress.function_declaration), (index, function));
        order.push((index, function));
        assert_eq!(producer.retained_statements(), source.body.len() - order.len());
    }
    assert_eq!(producer.remaining_units(), 0);
    assert_eq!(order, automatic_order);
    assert_same(&expected, &session.finish().unwrap());
    assert_eq!(format!("{source:?}"), original, "borrowed reference AST changed");
}

#[test]
fn borrowed_sessions_remain_borrowed_and_reject_ownership_transfer_without_poisoning() {
    let source = fixture();
    let original = format!("{source:?}");
    let expected = Compiler::with_profile(profile()).compile_with_imports(&source).unwrap();
    let mut session = Compiler::with_profile(profile()).begin_emission(&source).unwrap();
    assert!(session.take_owned_units().is_err());
    assert!(session.pending_chunks().is_ok());
    assert_same(&expected, &session.finish().unwrap());
    assert_eq!(format!("{source:?}"), original);
}

#[test]
fn owned_units_preserve_nonempty_host_import_metadata() {
    let mut source = fixture();
    source.imports = vec![
        Import { kind: ImportKind::Named { path: "vybe:prepared-test".into(),
            names: vec![ImportName { name: "exported".into(), alias: Some("local_export".into()) }],
            level: 0 }, span: Default::default() },
        Import { kind: ImportKind::Wildcard { path: "vybe:prepared-namespace".into(),
            alias: Some("namespace".into()) }, span: Default::default() },
    ];
    let expected = Compiler::with_profile(profile()).compile_with_imports(&source).unwrap();
    assert!(!expected.host_imports.named.is_empty());
    assert!(!expected.host_imports.wildcard.is_empty());
    let mut session = Compiler::with_profile(profile()).begin_emission_owned(source).unwrap();
    let mut producer = session.take_owned_units().unwrap();
    while let Some(unit) = producer.next_unit() { session.emit_unit(unit).unwrap(); }
    drop(producer);
    assert_same(&expected, &session.finish().unwrap());
}

#[test]
fn owned_units_preserve_partial_class_merge_and_statement_stream_handoff() {
    let mut source = fixture();
    for field in ["left", "right"] {
        source.body.push(Statement::new(StmtKind::ClassDecl {
            name: "PartialRecord".into(), parents: Vec::new(), interfaces: Vec::new(),
            members: vec![ClassMember::Field { name: field.into(), type_hint: None,
                init: Some(int(3)), modifiers: Default::default(), with_events: false,
                array_bounds: None, storage: None }],
            modifiers: ClassModifiers { is_partial: true, ..Default::default() },
            decorators: Vec::new(),
        }));
    }
    let expected = Compiler::with_profile(profile()).compile_with_imports(&source).unwrap();
    let mut session = Compiler::with_profile(profile()).begin_emission_owned(source.clone()).unwrap();
    let mut producer = session.take_owned_units().unwrap();
    while let Some(unit) = producer.next_unit() { session.emit_unit(unit).unwrap(); }
    drop(producer);
    assert_same(&expected, &session.finish().unwrap());

    let mut stream = Compiler::with_profile(profile()).begin_statement_emission(&source).unwrap();
    for (index, statement) in source.body.into_iter().enumerate() {
        stream.emit_statement(index, statement).unwrap();
    }
    assert_same(&expected, &stream.finish().unwrap());
}

#[test]
fn detached_units_cannot_be_skipped_reused_or_substituted_across_sessions() {
    let mut first = Compiler::with_profile(profile()).begin_emission_owned(fixture()).unwrap();
    let mut second = Compiler::with_profile(profile()).begin_emission_owned(fixture()).unwrap();
    let mut producer = first.take_owned_units().unwrap();
    let _second_producer = second.take_owned_units().unwrap();
    assert!(second.emit_unit(producer.next_unit().unwrap()).is_err());
    assert!(second.pending_chunks().is_err());
    assert!(second.emit_next().is_err());
    assert!(second.finish().is_err());
    assert!(first.emit_unit(producer.next_unit().unwrap()).is_err());
    assert!(first.finish().is_err());

    let mut incomplete = Compiler::with_profile(profile()).begin_emission_owned(fixture()).unwrap();
    let mut producer = incomplete.take_owned_units().unwrap();
    incomplete.emit_unit(producer.next_unit().unwrap()).unwrap();
    drop(producer);
    assert!(incomplete.finish().is_err(), "incomplete detached unit was finalized");
}

#[test]
fn late_emission_errors_match_borrowed_and_poison_the_owned_session() {
    let mut source = fixture();
    // Use an emitter rejection, not a frontend-only control-flow restriction.
    source.body.push(Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::ArrayTransform { op: ArrayTransformOp::PackMask,
            args: Vec::new(), order: ArrayTraversalOrder::RowMajor },
    ))));
    let expected = Compiler::with_profile(profile()).compile_with_imports(&source).unwrap_err();
    let mut session = Compiler::with_profile(profile()).begin_emission_owned(source).unwrap();
    let mut producer = session.take_owned_units().unwrap();
    let mut completed = 0;
    let actual = loop {
        let unit = producer.next_unit().expect("missing late error producer");
        match session.emit_unit(unit) {
            Ok(_) => completed += 1,
            Err(error) => break error,
        }
    };
    assert!(completed > 0);
    assert_eq!(actual, expected);
    assert!(session.pending_chunks().is_err());
    assert!(session.emit_next().is_err());
    assert!(session.finish().is_err());
}

#[test]
fn provider_transaction_commits_only_complete_successful_source_units() {
    let source = fixture();
    let expected = Compiler::with_profile(profile()).compile_with_imports(&source).unwrap();
    let emission = Compiler::with_profile(profile()).begin_statement_emission(&source).unwrap();
    let actual = emission.finish_with(|sink| {
        for (index, statement) in source.body.iter().cloned().enumerate() { sink(index, statement)?; }
        Ok(())
    }).unwrap();
    assert_same(&expected, &actual);

    for delivered in [1, source.body.len()] {
        let emission = Compiler::with_profile(profile()).begin_statement_emission(&source).unwrap();
        let error = emission.finish_with(|sink| {
            for (index, statement) in source.body.iter().take(delivered).cloned().enumerate() {
                sink(index, statement)?;
            }
            Err("provider rejected source after committed captures".into())
        }).unwrap_err();
        assert_eq!(error, "provider rejected source after committed captures");
    }

    let emission = Compiler::with_profile(profile()).begin_statement_emission(&source).unwrap();
    assert!(emission.finish_with(|sink| sink(0, source.body[0].clone()))
        .unwrap_err().contains("incomplete"));
}

#[test]
fn provider_transaction_rejects_duplicates_order_and_manifest_substitution() {
    let source = fixture();
    let emission = Compiler::with_profile(profile()).begin_statement_emission(&source).unwrap();
    assert!(emission.finish_with(|sink| sink(1, source.body[1].clone()))
        .unwrap_err().contains("out of order"));
    let emission = Compiler::with_profile(profile()).begin_statement_emission(&source).unwrap();
    assert!(emission.finish_with(|sink| {
        sink(0, source.body[0].clone())?;
        sink(0, source.body[0].clone())
    }).unwrap_err().contains("out of order"));
    let emission = Compiler::with_profile(profile()).begin_statement_emission(&source).unwrap();
    assert!(emission.finish_with(|sink| sink(0, Statement::new(StmtKind::Empty)))
        .unwrap_err().contains("differs from sealed normalized manifest"));
    let emission = Compiler::with_profile(profile()).begin_statement_emission(&source).unwrap();
    assert!(emission.finish_with(|sink| {
        let _ = sink(1, source.body[1].clone());
        Ok(())
    }).unwrap_err().contains("poisoned"), "provider concealed a sink rejection");
}
