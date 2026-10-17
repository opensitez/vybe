use vybe_ast::{BinOp, BitLane, ExprKind, Expression, IntOp, Lang, Module, Statement, StmtKind};
use vybe_compiler::primitives::platforms;
use vybe_runtime::{Chunk, VM};

fn compile(op: BinOp, left: Expression, right: Expression) -> Vec<Chunk> {
    let profile = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
    compile_profile(profile, op, left, right)
}

fn compile_profile(
    profile: vybe_runtime::profile::LanguageProfile,
    op: BinOp,
    left: Expression,
    right: Expression,
) -> Vec<Chunk> {
    compile_profile_directives(profile, op, left, right, Default::default())
}

fn compile_profile_directives(
    profile: vybe_runtime::profile::LanguageProfile,
    op: BinOp,
    left: Expression,
    right: Expression,
    directives: vybe_ast::Directives,
) -> Vec<Chunk> {
    let module = Module {
        name: "<integer-operand-literals>".into(),
        language: Lang::Unknown,
        body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
        ))))],
        imports: vec![],
        canon: Default::default(),
        directives,
    };
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}

fn static_word_profile() -> vybe_runtime::profile::LanguageProfile {
    vybe_compiler::profile::parse_profile(
        "[compiler]\n[builtin_types]\nint = [\"int\", \"integer\"]\n\
         [builtin_slots.int]\nadd = \"opcode:i32.add\"\nsub = \"opcode:i32.sub\"\n\
         mul = \"opcode:i32.mul\"\neq = \"opcode:i32.eq\"\nlt = \"opcode:i32.lt_s\"\n",
    )
    .unwrap()
}

#[test]
fn resolved_static_word_operators_keep_literal_words_native() {
    for (left, right) in [(0xffff_ffff, 1), (i64::from(i32::MAX), 1), (-7, 3)] {
        for (op, expected) in [
            (BinOp::Add, (left as i32).wrapping_add(right as i32)),
            (BinOp::Sub, (left as i32).wrapping_sub(right as i32)),
            (BinOp::Mul, (left as i32).wrapping_mul(right as i32)),
            (BinOp::Eq, i32::from(left as i32 == right as i32)),
            (BinOp::Lt, i32::from((left as i32) < (right as i32))),
        ] {
            let chunks = compile_profile(
                static_word_profile(),
                op,
                Expression::int(left),
                Expression::int(right),
            );
            let dump = vybe_runtime::debug::disassemble(&chunks[0]);
            assert!(!dump.contains("i64.trunc_f64_s"), "{dump}");
            assert!(!dump.contains("i32.wrap_i64"), "{dump}");
            assert_eq!(validate_and_run(chunks).as_i32(), expected);
        }
    }
}

#[test]
fn resolved_static_word_operators_retain_wider_literal_normalization() {
    for value in [
        i64::from(i32::MIN) - 1,
        i64::from(u32::MAX) + 1,
        9_007_199_254_740_993,
    ] {
        let chunks = compile_profile(
            static_word_profile(),
            BinOp::Add,
            Expression::int(value),
            Expression::int(0),
        );
        let dump = vybe_runtime::debug::disassemble(&chunks[0]);
        assert!(dump.contains("i64.trunc_f64_s"), "{dump}");
        assert!(dump.contains("i32.wrap_i64"), "{dump}");
        assert_eq!(
            validate_and_run(chunks).as_i32(),
            (value as f64) as i64 as i32
        );
    }
}

#[test]
fn resolved_static_word_loads_require_an_actual_producer_report() {
    use vybe_ast::Argument;
    use vybe_runtime::profile::{BuiltinDef, BuiltinEmit};
    let mut profile = static_word_profile();
    profile
        .builtin_return_types
        .insert("read_word".into(), "integer".into());
    profile.builtins.insert(
        "read_word".into(),
        BuiltinDef {
            emit: BuiltinEmit::Common("pointers.linear_i32_load8_u".into()),
            discarded_emit: None,
            min_args: 1,
            max_args: 1,
            i32_condition: false,
            slot: None,
        },
    );
    let read = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("read_word")),
        args: vec![Argument::positional(Expression::int(0))],
        optional: false,
    });
    for op in [BinOp::Add, BinOp::Eq, BinOp::Lt] {
        for (left, right) in [
            (read.clone(), Expression::int(1)),
            (Expression::int(1), read.clone()),
        ] {
            let chunks = compile_profile(profile.clone(), op.clone(), left, right);
            let dump = vybe_runtime::debug::disassemble(&chunks[0]);
            assert!(dump.contains("i32.load8_u"), "{dump}");
            assert!(!dump.contains("i64.trunc_f64_s"), "{dump}");
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm_with_options(
                    &chunks,
                    vybe_platform_wasm::WasmWriteOptions {
                        include_vybe_metadata: false,
                        include_name_section: true,
                    },
                ))
                .unwrap();
        }
    }
    profile.builtins.get_mut("read_word").unwrap().emit =
        BuiltinEmit::HostCall("test:word".into(), "read".into());
    let chunks = compile_profile(profile, BinOp::Add, read, Expression::int(1));
    assert!(vybe_runtime::debug::disassemble(&chunks[0]).contains("i64.trunc_f64_s"));
}

#[test]
fn dynamic_operator_policy_keeps_the_number_backed_path() {
    let chunks = compile_profile_directives(
        static_word_profile(),
        BinOp::Add,
        Expression::int(7),
        Expression::int(3),
        vybe_ast::Directives {
            type_resolution: Some(vybe_ast::TypeResolution::Dynamic),
            operator_dispatch: Some(vybe_ast::OperatorDispatch::RuntimeProtocol),
            ..Default::default()
        },
    );
    let dump = vybe_runtime::debug::disassemble(&chunks[0]);
    assert!(dump.contains("i64.trunc_f64_s"), "{dump}");
    assert_eq!(validate_and_run(chunks).as_i32(), 10);
}

#[test]
fn resolved_linear_word_results_skip_only_redundant_conversions() {
    use vybe_ast::Argument;
    use vybe_runtime::profile::{BuiltinDef, BuiltinEmit};
    for primitive in [
        "linear_i32_load",
        "linear_i32_load8_u",
        "linear_i32_load8_s",
        "linear_i32_load16_u",
        "linear_i32_load16_s",
    ] {
        let mut profile = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
        profile.builtins.insert(
            "read_word".into(),
            BuiltinDef {
                emit: BuiltinEmit::Common(format!("pointers.{primitive}")),
                discarded_emit: None,
                min_args: 1,
                max_args: 1,
                i32_condition: false,
                slot: None,
            },
        );
        let read = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("read_word")),
            args: vec![Argument::positional(Expression::int(0))],
            optional: false,
        });
        for (left, right) in [
            (read.clone(), Expression::int(1)),
            (Expression::int(1), read.clone()),
        ] {
            let chunks = compile_profile(
                profile.clone(),
                BinOp::Integer(IntOp::Add, BitLane::W32),
                left,
                right,
            );
            let dump = vybe_runtime::debug::disassemble(&chunks[0]);
            assert!(dump.contains("i32.load"), "{dump}");
            assert!(!dump.contains("i64.trunc_f64_s"), "{dump}");
            assert!(!dump.contains("i32.wrap_i64"), "{dump}");
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm_with_options(
                    &chunks,
                    vybe_platform_wasm::WasmWriteOptions {
                        include_vybe_metadata: false,
                        include_name_section: true,
                    },
                ))
                .unwrap();
        }
        // The same source spelling resolved to a host call supplies no proof.
        profile.builtins.get_mut("read_word").unwrap().emit =
            BuiltinEmit::HostCall("test:word".into(), "read".into());
        let chunks = compile_profile(
            profile,
            BinOp::Integer(IntOp::Add, BitLane::W32),
            read,
            Expression::int(1),
        );
        assert!(vybe_runtime::debug::disassemble(&chunks[0]).contains("i64.trunc_f64_s"));
    }
}

#[test]
fn word_request_reports_actual_owned_emission_without_changing_value_emission() {
    use vybe_compiler::primitives::dispatch::emit_common_with_result;
    use vybe_runtime::registry::emitter_results::{EmitOutcome, EmitRequest, ResultRepresentation};
    let mut ordinary = vec![Chunk::new("ordinary")];
    let mut word = vec![Chunk::new("word")];
    assert_eq!(
        emit_common_with_result(
            "pointers.linear_i32_load8_u",
            &mut ordinary,
            0,
            1,
            0,
            EmitRequest::Value
        ),
        EmitOutcome::Handled(ResultRepresentation::Unknown)
    );
    assert_eq!(
        emit_common_with_result(
            "pointers.linear_i32_load8_u",
            &mut word,
            0,
            1,
            0,
            EmitRequest::I32Word
        ),
        EmitOutcome::Handled(ResultRepresentation::RawI32Word)
    );
    assert_eq!(ordinary[0].code, word[0].code);
    let before = word[0].code.clone();
    assert_eq!(
        emit_common_with_result(
            "pointers.not_a_load",
            &mut word,
            0,
            1,
            0,
            EmitRequest::I32Word
        ),
        EmitOutcome::Unhandled
    );
    assert_eq!(word[0].code, before);
}

fn validate_and_run(chunks: Vec<Chunk>) -> vybe_runtime::Value {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm_with_options(
            &chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            },
        ))
        .unwrap();
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap()
}

#[test]
fn w32_word_literals_skip_normalization_in_either_operand() {
    for value in [
        i64::from(i32::MIN),
        -1,
        0,
        i64::from(i32::MAX),
        0x8000_0000,
        0xffff_ffff,
    ] {
        for (left, right) in [
            (value, 0),
            (0, value),
            (value, 1),
            (1, value),
            (value, value),
        ] {
            for (op, expected) in [
                (IntOp::Eq, i32::from(left as i32 == right as i32)),
                (IntOp::LtS, i32::from((left as i32) < (right as i32))),
                (IntOp::LtU, i32::from((left as u32) < (right as u32))),
                (IntOp::Add, (left as i32).wrapping_add(right as i32)),
                (IntOp::Sub, (left as i32).wrapping_sub(right as i32)),
                (IntOp::Mul, (left as i32).wrapping_mul(right as i32)),
            ] {
                let chunks = compile(
                    BinOp::Integer(op, BitLane::W32),
                    Expression::int(left),
                    Expression::int(right),
                );
                let dump = vybe_runtime::debug::disassemble(&chunks[0]);
                assert!(!dump.contains("i64.trunc_f64_s"), "{dump}");
                assert!(!dump.contains("i32.wrap_i64"), "{dump}");
                assert_eq!(
                    validate_and_run(chunks).as_i32(),
                    expected,
                    "{op:?}: {left}, {right}"
                );
            }
        }
    }
}

#[test]
fn larger_literals_and_float_operands_retain_numeric_normalization() {
    for left in [
        Expression::int(i64::from(i32::MIN) - 1),
        Expression::int(i64::from(u32::MAX) + 1),
        Expression::int(9_007_199_254_740_993),
        Expression::float(4_294_967_295.75),
    ] {
        let expected = match &left.kind {
            ExprKind::Lit(vybe_ast::Literal::Int(n)) => (*n as f64) as i64 as i32,
            ExprKind::Lit(vybe_ast::Literal::Float(n)) => *n as i64 as i32,
            _ => unreachable!(),
        };
        let chunks = compile(
            BinOp::Integer(IntOp::Add, BitLane::W32),
            left,
            Expression::int(0),
        );
        let dump = vybe_runtime::debug::disassemble(&chunks[0]);
        assert!(dump.contains("i64.trunc_f64_s"), "{dump}");
        assert!(dump.contains("i32.wrap_i64"), "{dump}");
        assert_eq!(validate_and_run(chunks).as_i32(), expected);
    }
}

#[test]
fn w64_and_generic_arithmetic_keep_positive_word_literals() {
    for op in [BinOp::Integer(IntOp::Add, BitLane::W64), BinOp::Add] {
        let chunks = compile(op, Expression::int(0xffff_ffff), Expression::int(1));
        assert_eq!(validate_and_run(chunks).as_f64(), 4_294_967_296.0);
    }
}
