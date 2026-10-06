use vybe_compiler::primitives::{bundle, globals, math, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/erf_optimization_baseline.rs"]
mod baseline;

fn program(old: bool, complement: bool, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<erf-optimization>")];
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "input", 37);
        match (old, complement) {
            (true, false) => baseline::emit_erf(&mut chunks[0], 37),
            (false, false) => math::emit_erf(&mut chunks[0], 37),
            (true, true) => baseline::emit_erfc(&mut chunks[0], 37),
            (false, true) => math::emit_erfc(&mut chunks[0], 37),
        }
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 37);
        }
    }
    chunks[0].emit_op(Op::RETURN, 37);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn runtime(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn assert_exact_value(before: Value, after: Value) {
    match (before, after) {
        (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        (Value::F32(a), Value::F32(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        (a, b) => assert_eq!(a, b),
    }
}

#[test]
fn shared_erf_preserves_ieee_bits_approximation_and_complement() {
    for complement in [false, true] {
        let old = program(true, complement, 1);
        let new = program(false, complement, 1);
        validate(&new);
        assert!(
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(&old))
                .is_err(),
            "captured erf has void branches carrying f64 results"
        );
        let mut inputs: Vec<_> = (-200..=200).map(|i| Value::F64(i as f64 / 16.0)).collect();
        inputs.extend([
            Value::F64(-0.0),
            Value::F64(0.0),
            Value::F64(f64::INFINITY),
            Value::F64(f64::NEG_INFINITY),
            Value::F64(f64::NAN),
            Value::F64(f64::from_bits(0x7ff8_0000_0000_0011)),
            Value::F64(f64::from_bits(0xfff8_0000_0000_0011)),
            Value::F64(f64::from_bits(1)),
            Value::F64(f64::from_bits(0x8000_0000_0000_0001)),
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::Bool(true),
            Value::I32(-1),
            Value::I32(1),
            Value::String("1.25".into()),
            Value::String("not a number".into()),
            Value::F32(f32::NAN),
            Value::F32(-0.0),
        ]);
        for input in inputs {
            let before = runtime(input.clone()).run(old.clone());
            let after = runtime(input.clone()).run(new.clone());
            match before {
                Ok(value) => assert_exact_value(value, after.unwrap()),
                Err(before) => {
                    let after = after.unwrap_err();
                    assert_eq!(before.message, after.message);
                    assert_eq!(before.is_trap(), after.is_trap());
                    // A shared body adds its own frame. The user call site
                    // still identifies the same caller chunk and source line.
                    let caller = |error: &vybe_runtime::VMError| {
                        error
                            .call_stack
                            .iter()
                            .filter(|frame| frame.chunk_name == "<erf-optimization>")
                            .map(|frame| (frame.chunk_name.clone(), frame.line))
                            .collect::<Vec<_>>()
                    };
                    assert_eq!(caller(&before), caller(&after));
                }
            }
        }
        for (input, expected) in [
            (0.0, if complement { 1.0 } else { 0.0 }),
            (f64::INFINITY, if complement { 0.0 } else { 1.0 }),
            (f64::NEG_INFINITY, if complement { 2.0 } else { -1.0 }),
        ] {
            assert_eq!(
                runtime(Value::F64(input)).run(new.clone()).unwrap(),
                Value::F64(expected)
            );
        }
    }
}

#[test]
fn shared_erf_has_one_body_fewer_locals_and_idempotent_finalization() {
    for complement in [false, true] {
        let old = program(true, complement, 20);
        let mut new = program(false, complement, 20);
        validate(&new);
        let bodies: Vec<_> = new
            .iter()
            .filter(|chunk| chunk.name == "__stdlib_erf")
            .collect();
        assert_eq!(bodies.len(), 1);
        assert_eq!(bodies[0].arity, 1);
        assert_eq!(bodies[0].local_count, 3);
        let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
        assert!(bytes(&new) < bytes(&old));
        assert!(
            vybe_platform_wasm::write_wasm(&new).len() < vybe_platform_wasm::write_wasm(&old).len()
        );
        let before = vybe_platform_wasm::write_wasm(&new);
        let code_before = bytes(&new);
        bundle::finalize_with_runtime_helpers(&mut new);
        assert_eq!(bytes(&new), code_before);
        assert_eq!(vybe_platform_wasm::write_wasm(&new), before);
        validate(&new);
        println!(
            "erf complement={complement} 20 calls: bytecode={}/{} wasm={}/{}",
            bytes(&old),
            bytes(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            before.len()
        );
    }
}

#[test]
#[ignore = "interleaved opt-level-0 erf execution benchmark"]
fn shared_erf_execution_benchmark() {
    for complement in [false, true] {
        let programs = [
            program(true, complement, 256),
            program(false, complement, 256),
        ];
        let mut times = [vec![], vec![]];
        let mut results = [Value::Null, Value::Null];
        for round in 0..9 {
            for which in [round % 2, 1 - round % 2] {
                let mut vm = runtime(Value::F64(1.25));
                let chunks = programs[which].clone();
                let start = std::time::Instant::now();
                results[which] = vm.run(chunks).unwrap();
                times[which].push(start.elapsed());
            }
        }
        assert_eq!(results[0], results[1]);
        for samples in &mut times {
            samples.sort();
        }
        println!(
            "erf complement={complement} 256 calls: before={}us after={}us speedup={:.2}x",
            times[0][4].as_micros(),
            times[1][4].as_micros(),
            times[0][4].as_secs_f64() / times[1][4].as_secs_f64()
        );
    }
}

fn source_collision_program(alias_helper_global: bool, callable_kind: u8) -> Vec<Chunk> {
    source_collision_program_padded(alias_helper_global, callable_kind, 1)
}

fn boundary_placeholder() -> Chunk {
    let mut chunk = Chunk::new("__stdlib_unused_boundary");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 1);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn source_collision_program_padded(
    alias_helper_global: bool,
    callable_kind: u8,
    initial_chunks: usize,
) -> Vec<Chunk> {
    use vybe_ast::{
        Argument, BindingPattern, Directives, ExprKind, Expression, Lang, Literal, Module, Param,
        PassBy, Statement, StmtKind, VarDeclKind, VarDeclarator,
    };
    // Common AST corresponding to legal Python source:
    // def __stdlib_erf(x): return 42.0
    // [__vybe_erf = __stdlib_erf]
    // answer = math.erf(1.25)
    // A synthetic final return observes answer for the test harness.
    // Use the production Python declaration profile and compiler; no fake
    // helper dispatcher or helper registry is installed for this source test.
    let mut body = vec![Statement::new(StmtKind::FunctionDecl {
        name: "__stdlib_erf".into(),
        params: vec![Param {
            name: "x".into(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }],
        return_type: None,
        body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
            ExprKind::Lit(Literal::Float(42.0)),
        ))))],
        modifiers: Default::default(),
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    })];
    if alias_helper_global {
        body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__vybe_erf".into()),
                type_hint: None,
                init: Some(Expression::ident("__stdlib_erf")),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }));
    }
    let math_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__py_math_erf")),
        args: vec![Argument::positional(Expression::new(ExprKind::Lit(
            Literal::Float(1.25),
        )))],
        optional: false,
    });
    if alias_helper_global {
        body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("alias_answer".into()),
                type_hint: None,
                init: Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__vybe_erf")),
                    args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                        Literal::Float(1.25),
                    )))],
                    optional: false,
                })),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }));
    }
    if callable_kind == 1 {
        body.push(Statement::new(StmtKind::FunctionDecl {
            name: "calculate".into(),
            params: vec![],
            return_type: None,
            body: vec![Statement::new(StmtKind::Return(Some(math_call.clone())))],
            modifiers: Default::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        }));
    } else if callable_kind == 2 {
        body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("calculate".into()),
                type_hint: None,
                init: Some(Expression::new(ExprKind::Lambda {
                    params: vec![],
                    body: vybe_ast::LambdaBody::Expr(Box::new(math_call.clone())),
                    is_async: false,
                    captures: vec![],
                })),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }));
    }
    let answer = if callable_kind == 0 {
        math_call
    } else {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("calculate")),
            args: vec![],
            optional: false,
        })
    };
    body.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("answer".into()),
            type_hint: None,
            init: Some(answer),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Var,
    }));
    body.push(Statement::new(StmtKind::Return(Some(Expression::ident(
        "answer",
    )))));
    // Use only the public compiler API. Discarded, capture-free Python
    // lambdas consume real source function chunks without adding root names
    // or accumulating declaration/closure scratch slots.
    if initial_chunks > 1 {
        let mut prefix = Vec::with_capacity(initial_chunks - 1 + body.len());
        for _ in 1..initial_chunks {
            prefix.push(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Lambda {
                    params: vec![],
                    body: vybe_ast::LambdaBody::Expr(Box::new(Expression::new(ExprKind::Lit(
                        Literal::Null,
                    )))),
                    is_async: false,
                    captures: vec![],
                },
            ))));
        }
        prefix.extend(body);
        body = prefix;
    }
    let module = Module {
        name: "<erf-source-collision>".into(),
        language: Lang::Python,
        body,
        imports: vec![],
        directives: Directives::default(),
        canon: Default::default(),
    };
    let profile = vybe_compiler::profile::parse_profile(include_str!(
        "../../../languages/python/src/profile"
    ))
    .unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}

#[test]
fn ordinary_source_function_named_like_erf_helper_keeps_generated_body() {
    let chunks = source_collision_program(false, 0);
    validate(&chunks);
    assert!(
        !chunks[0]
            .global_inits
            .iter()
            .any(|init| init.name == "__vybe_erf")
    );
    let helper = chunks
        .iter()
        .find(|chunk| {
            chunk.local_names.iter().any(|local| {
                local.origin == vybe_runtime::chunk::LocalOrigin::Compiler
                    && local.name == "<private-helper-body:__stdlib_erf>"
            })
        })
        .expect("compiler installs its own private erf body");
    assert_eq!(helper.name, "__stdlib_erf");
    assert_eq!(helper.local_count, 3);
    assert_eq!(
        chunks
            .iter()
            .filter(|chunk| chunk.name == "__stdlib_erf")
            .count(),
        2,
        "source name alone must not suppress the generated algorithm"
    );
    let expected = runtime(Value::F64(1.25))
        .run(program(true, false, 1))
        .unwrap();
    assert_exact_value(expected, runtime(Value::Null).run(chunks).unwrap());
}

#[test]
fn explicit_source_helper_global_alias_preserves_both_source_and_math_results() {
    let expected = runtime(Value::F64(1.25))
        .run(program(true, false, 1))
        .unwrap();
    for callable_kind in 0..=2 {
        let mut chunks = source_collision_program(true, callable_kind);
        validate(&chunks);
        let bytes: Vec<_> = chunks.iter().map(|chunk| chunk.code.clone()).collect();
        let count = chunks.len();
        bundle::finalize_with_runtime_helpers(&mut chunks);
        assert_eq!(chunks.len(), count);
        assert_eq!(
            chunks
                .iter()
                .map(|chunk| chunk.code.clone())
                .collect::<Vec<_>>(),
            bytes
        );
        let mut vm = runtime(Value::Null);
        assert_exact_value(expected.clone(), vm.run(chunks).unwrap());
        assert_eq!(vm.global("alias_answer"), Some(&Value::F64(42.0)));
    }
}

#[test]
fn private_erf_fixup_ignores_source_markers_and_ordinary_max_index_refs() {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    let mut caller = Chunk::new("<erf-provenance>");
    caller.emit_f64_const(1.25, 1);
    math::emit_erf(&mut caller, 1);
    let pending = caller
        .local_names
        .iter()
        .find(|local| local.name.starts_with("<private-helper-ref:"))
        .unwrap()
        .clone();
    // A duplicated compiler record is harmless; a source-origin lookalike
    // never authorizes a fixup even when the real operand is the sentinel.
    caller.local_names.push(pending);
    caller.emit_op(Op::RETURN, 1);
    let mut source = Chunk::new("<source-ref>");
    vybe_compiler::primitives::functions::emit_ref_func(&mut source, u16::MAX as usize, 0, 1);
    source.local_names.push(LocalName::new(
        0,
        "<private-helper-ref:__stdlib_erf@0>",
        LocalOrigin::Source,
    ));
    let mut chunks = vec![caller, source];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    assert_eq!(chunks[1].read_u16(4), u16::MAX);
    assert_eq!(chunks[1].local_names[0].origin, LocalOrigin::Source);
    assert!(
        !chunks[0]
            .local_names
            .iter()
            .any(|local| local.origin == LocalOrigin::Compiler
                && local.name.starts_with("<private-helper-ref:"))
    );

    // A newly generated caller after initial finalization reuses the existing
    // private body. This covers incremental helper discovery and finalization.
    let previous_len = chunks.len();
    let mut next = Chunk::new("<next-erf-caller>");
    next.emit_f64_const(1.25, 1);
    math::emit_erf(&mut next, 1);
    next.emit_op(Op::RETURN, 1);
    chunks.push(next);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    assert_eq!(chunks.len(), previous_len + 1);
    assert_eq!(
        chunks
            .iter()
            .filter(|chunk| chunk.name == "__stdlib_erf")
            .count(),
        1
    );
}

#[test]
fn private_erf_exact_u16_boundary_inlines_when_helper_would_overflow() {
    let expected = runtime(Value::F64(1.25))
        .run(program(true, false, 1))
        .unwrap();
    for initial_chunks in [u16::MAX as usize, u16::MAX as usize + 1] {
        let mut caller = Chunk::new("<erf-function-limit>");
        caller.emit_f64_const(1.25, 1);
        math::emit_erf(&mut caller, 1);
        caller.emit_op(Op::RETURN, 1);
        let mut chunks = vec![caller];
        chunks.resize_with(initial_chunks, boundary_placeholder);
        bundle::finalize_with_runtime_helpers(&mut chunks);
        let fits = initial_chunks == u16::MAX as usize;
        assert_eq!(chunks.len(), initial_chunks + usize::from(fits));
        assert_eq!(
            chunks
                .iter()
                .filter(|chunk| chunk
                    .local_names
                    .iter()
                    .any(|local| local.name == "<private-helper-body:__stdlib_erf>"))
                .count(),
            usize::from(fits)
        );
        assert_eq!(chunks[0].code.len(), chunks[0].lines.len());
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        validate(&chunks);
        let bytes = chunks[0].code.clone();
        bundle::finalize_with_runtime_helpers(&mut chunks);
        assert_eq!(chunks[0].code, bytes);
        assert_exact_value(expected.clone(), runtime(Value::Null).run(chunks).unwrap());
    }
}

#[test]
fn oversized_source_erf_keeps_alias_function_and_lambda_semantics() {
    let expected = runtime(Value::F64(1.25))
        .run(program(true, false, 1))
        .unwrap();
    for callable_kind in 0..=2 {
        let unpadded = source_collision_program(true, callable_kind);
        let fixed_chunks = unpadded
            .iter()
            .filter(|chunk| {
                !chunk.local_names.iter().any(|local| {
                    local.origin == vybe_runtime::chunk::LocalOrigin::Compiler
                        && local.name == "<private-helper-body:__stdlib_erf>"
                })
            })
            .count();
        // Padding contributes initial_chunks - 1 actual lambda chunks. Derive
        // the other chunks from real compilation, including closure factories.
        let initial_chunks = u16::MAX as usize + 2 - fixed_chunks;
        let chunks = source_collision_program_padded(true, callable_kind, initial_chunks);
        assert_eq!(chunks.len(), u16::MAX as usize + 1);
        assert!(!chunks.iter().any(|chunk| {
            chunk
                .local_names
                .iter()
                .any(|local| local.name == "<private-helper-body:__stdlib_erf>")
        }));
        validate(&chunks);
        let mut vm = runtime(Value::Null);
        assert_exact_value(expected.clone(), vm.run(chunks).unwrap());
        assert_eq!(vm.global("alias_answer"), Some(&Value::F64(42.0)));
    }
}

#[test]
fn public_helper_exclusions_do_not_remove_required_private_erf_body() {
    let mut caller = Chunk::new("<excluded-erf>");
    caller.emit_f64_const(1.25, 1);
    math::emit_erf(&mut caller, 1);
    caller.emit_op(Op::RETURN, 1);
    let mut chunks = vec![caller];
    bundle::finalize_with_runtime_helpers_excluding(&mut chunks, &["__stdlib_erf"]);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate(&chunks);
    assert!(
        !chunks[0]
            .global_inits
            .iter()
            .any(|init| init.name == "__vybe_erf")
    );
    let expected = runtime(Value::F64(1.25))
        .run(program(true, false, 1))
        .unwrap();
    assert_exact_value(expected, runtime(Value::Null).run(chunks).unwrap());
}

#[test]
fn private_erf_refs_use_normal_loader_relocation_after_prior_programs() {
    let expected = runtime(Value::F64(1.25))
        .run(program(true, false, 1))
        .unwrap();
    let mut vm = runtime(Value::F64(1.25));
    vm.run(program(true, false, 1)).unwrap();
    for callable_kind in 0..=2 {
        assert_exact_value(
            expected.clone(),
            vm.run(source_collision_program(true, callable_kind))
                .unwrap(),
        );
        assert_exact_value(expected.clone(), vm.run(program(false, false, 20)).unwrap());
    }
}

#[test]
fn oversized_mixed_private_calls_relocate_metadata_and_real_generator_handlers() {
    use vybe_compiler::primitives::{functions, generators, random};
    use vybe_runtime::chunk::StackSwitchHandler;
    let mut chunks = vec![Chunk::new("<oversized-generator-caller>")];
    random::emit_next_unit(&mut chunks, 0, 7);
    chunks[0].emit_op(Op::DROP, 7);
    chunks[0].emit_f64_const(1.25, 8);
    math::emit_erf(&mut chunks[0], 8);
    chunks[0].emit_op(Op::DROP, 8);
    functions::emit_ref_func(&mut chunks[0], 1, 0, 9);
    chunks[0].emit_op(Op::CONT_NEW, 9);
    generators::emit_next(&mut chunks[0], 9);
    chunks[0].emit_op(Op::DROP, 9); // has_more
    chunks[0].emit_op(Op::RETURN, 9); // yielded value
    let (resume, forward_target) = chunks[0]
        .stack_switch_handlers
        .iter()
        .next()
        .map(|(&key, handlers)| (key, handlers[0].label_index))
        .unwrap();
    chunks[0].user_code_offset = Some(resume as u32);

    // Unreachable standard-instruction contexts still carry offset metadata.
    let block = chunks[0].code.len();
    chunks[0].emit_block_i32(10);
    chunks[0].emit_i32_const(0, 10);
    chunks[0].emit_end(10);
    chunks[0].emit_op(Op::DROP, 10);
    let indirect = chunks[0].code.len();
    chunks[0].emit_op(Op::CALL_INDIRECT, 10);
    for byte in [0, 0, 1] {
        chunks[0].emit(byte, 10);
    }
    chunks[0]
        .call_indirect_sigs
        .insert(indirect, "->externref".into());
    chunks[0]
        .call_indirect_canon
        .insert(indirect, "fixture:fn-type".into());
    let structural_resume = chunks[0].code.len();
    generators::emit_resume(&mut chunks[0], 10);
    chunks[0].stack_switch_handlers.insert(
        structural_resume,
        vec![StackSwitchHandler {
            kind: 0,
            tag_index: 0,
            label_index: 0,
        }],
    );
    chunks[0].emit_op(Op::RETURN, 10);
    let original_len = chunks[0].code.len();

    let mut generator = functions::create_function_chunk("<yielding-body>", 0);
    generator.emit_f64_const(17.0, 11);
    generators::emit_suspend(&mut generator, 11);
    generator.emit_op(Op::DROP, 11);
    generator.emit_f64_const(19.0, 11);
    generator.emit_op(Op::RETURN, 11);
    chunks.push(generator);
    chunks.resize_with(u16::MAX as usize + 1, boundary_placeholder);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    assert_eq!(
        chunks.len(),
        u16::MAX as usize + 1,
        "no orphan private body"
    );
    let delta = chunks[0].code.len() - original_len;
    assert!(delta > 0);
    assert_eq!(chunks[0].user_code_offset, Some((resume + delta) as u32));
    assert!(chunks[0].block_i32_results.contains(&(block + delta)));
    assert_eq!(
        chunks[0]
            .call_indirect_sigs
            .get(&(indirect + delta))
            .map(String::as_str),
        Some("->externref")
    );
    assert_eq!(
        chunks[0]
            .call_indirect_canon
            .get(&(indirect + delta))
            .map(String::as_str),
        Some("fixture:fn-type")
    );
    assert_eq!(
        chunks[0].stack_switch_handlers[&(resume + delta)][0].label_index,
        forward_target + delta as u32
    );
    assert_eq!(
        chunks[0].stack_switch_handlers[&(structural_resume + delta)][0].label_index,
        0
    );
    assert_eq!(chunks[0].code.len(), chunks[0].lines.len());
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    // The source generator producer stores VM byte offsets in its handler,
    // while the writer expects standard label depths. Its existing native gap
    // is independent of private inlining; direct primitive boundary modules
    // above remain fully validated. Here the real VM exercises the handler.
    let mut vm = runtime(Value::Null);
    vm.set_global("__vybe_rng", Value::I32(42));
    assert_eq!(vm.run(chunks).unwrap(), Value::F64(17.0));
    let mut state = 42i32;
    state ^= state.wrapping_shl(13);
    state ^= ((state as u32) >> 17) as i32;
    state ^= state.wrapping_shl(5);
    assert_eq!(vm.global("__vybe_rng"), Some(&Value::I32(state)));
}

#[test]
fn mixed_private_boundary_keeps_fitting_ref_and_inlines_only_overflowing_body() {
    use vybe_compiler::primitives::random;
    let expected_erf = runtime(Value::F64(1.25))
        .run(program(true, false, 1))
        .unwrap();
    let mut chunks = vec![Chunk::new("<mixed-private-boundary>")];
    // Random occurs first, while deterministic helper ordering places Erf at
    // the last encodable index. Descending fixups must preserve the patched
    // Erf reference when the preceding Random call expands inline.
    random::emit_next_unit(&mut chunks, 0, 1);
    globals::emit_write(&mut chunks[0], "draw", 1);
    chunks[0].emit_f64_const(1.25, 1);
    math::emit_erf(&mut chunks[0], 1);
    chunks[0].emit_op(Op::RETURN, 1);
    chunks.resize_with(u16::MAX as usize, boundary_placeholder);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    assert_eq!(chunks.len(), u16::MAX as usize + 1);
    assert_eq!(chunks[u16::MAX as usize].name, "__stdlib_erf");
    assert!(
        !chunks
            .iter()
            .any(|chunk| chunk.name == "__stdlib_random_next_unit")
    );
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate(&chunks);
    let mut vm = runtime(Value::Null);
    vm.set_global("__vybe_rng", Value::I32(42));
    assert_exact_value(expected_erf, vm.run(chunks).unwrap());
    let mut state = 42i32;
    state ^= state.wrapping_shl(13);
    state ^= ((state as u32) >> 17) as i32;
    state ^= state.wrapping_shl(5);
    assert_eq!(vm.global("__vybe_rng"), Some(&Value::I32(state)));
    assert_eq!(
        vm.global("draw"),
        Some(&Value::F64((state & 0x7fffff) as f64 / 8388608.0))
    );
}

// Append to erf_optimization_emission_perf.rs after authorizing the runtime patch.
#[test]
fn small_private_erf_after_large_prior_load_preserves_source_aliases_and_delayed_calls() {
    let expected = runtime(Value::F64(1.25))
        .run(program(true, false, 1))
        .unwrap();
    for kind in 0..=2 {
        let mut vm = runtime(Value::Null);
        vm.run(vec![boundary_placeholder(); u16::MAX as usize + 1])
            .unwrap();
        let chunks = source_collision_program(true, kind);
        validate(&chunks); // Still validates the original small native module.
        assert!(chunks.len() < 16);
        assert_exact_value(expected.clone(), vm.run(chunks).unwrap());
        assert_eq!(vm.global("alias_answer"), Some(&Value::F64(42.0)));
        let retained = if kind == 1 {
            Some(
                vm.global("calculate")
                    .expect("real Source function")
                    .clone(),
            )
        } else {
            None
        };
        // Loading later code must not redirect private calls in retained bodies.
        vm.run(vec![boundary_placeholder(); 3]).unwrap();
        if let Some(retained) = retained {
            assert_exact_value(expected.clone(), vm.invoke_callback(&retained, &[]));
        }
    }
}

fn standard_ref_module(answer: u8, padding_functions: usize) -> Vec<u8> {
    fn leb(out: &mut Vec<u8>, mut n: usize) {
        loop {
            let b = (n & 127) as u8;
            n >>= 7;
            out.push(b | if n != 0 { 128 } else { 0 });
            if n == 0 {
                break;
            }
        }
    }
    fn section(out: &mut Vec<u8>, id: u8, payload: Vec<u8>) {
        out.push(id);
        leb(out, payload.len());
        out.extend(payload);
    }
    let mut bytes = b"\0asm\x01\x00\x00\x00".to_vec();
    if padding_functions != 0 {
        section(&mut bytes, 1, vec![1, 0x60, 0, 1, 0x7f]);
        let mut funcs = Vec::new();
        leb(&mut funcs, padding_functions);
        funcs.resize(funcs.len() + padding_functions, 0);
        section(&mut bytes, 3, funcs);
        let mut code = Vec::new();
        leb(&mut code, padding_functions);
        for _ in 0..padding_functions {
            code.extend([4, 0, 0x41, 0, 0x0b]);
        }
        section(&mut bytes, 10, code);
    } else {
        // Standard Wasm only: target ()->i32, _start ()->(), one mutable
        // funcref global. _start materializes target with ref.func and saves it.
        section(&mut bytes, 1, vec![2, 0x60, 0, 1, 0x7f, 0x60, 0, 0]);
        section(&mut bytes, 3, vec![2, 0, 1]);
        section(&mut bytes, 6, vec![1, 0x70, 1, 0xd0, 0x70, 0x0b]);
        section(
            &mut bytes,
            7,
            vec![1, 6, b'_', b's', b't', b'a', b'r', b't', 0, 1],
        );
        // Declare target for standard ref.func validation.
        section(&mut bytes, 9, vec![1, 3, 0, 1, 0]);
        section(
            &mut bytes,
            10,
            vec![2, 4, 0, 0x41, answer, 0x0b, 6, 0, 0xd2, 0, 0x24, 0, 0x0b],
        );
    }
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    bytes
}

#[test]
fn standard_decoded_wasm_ref_func_keeps_module_target_after_large_prior_load_and_later_load() {
    for linked in [false, true] {
        let mut vm = VM::new();
        let padding = vybe_platform_wasm::read_wasm(&standard_ref_module(0, 65536)).unwrap();
        assert_eq!(padding.len(), 65537);
        if linked {
            vm.run_linked(padding, vec![]).unwrap();
        } else {
            vm.run(padding).unwrap();
        }
        let first = vybe_platform_wasm::read_wasm(&standard_ref_module(7, 0)).unwrap();
        if linked {
            vm.run_linked(first, vec![]).unwrap();
        } else {
            vm.run(first).unwrap();
        }
        let retained = vm
            .global("__wasm_global_0")
            .expect("standard global.set result")
            .clone();
        assert_eq!(vm.invoke_callback(&retained, &[]), Value::I32(7));
        let later = vybe_platform_wasm::read_wasm(&standard_ref_module(9, 0)).unwrap();
        if linked {
            vm.run_linked(later, vec![]).unwrap();
        } else {
            vm.run(later).unwrap();
        }
        let current = vm.global("__wasm_global_0").unwrap().clone();
        assert_eq!(vm.invoke_callback(&current, &[]), Value::I32(9));
        assert_eq!(vm.invoke_callback(&retained, &[]), Value::I32(7));
    }
}
