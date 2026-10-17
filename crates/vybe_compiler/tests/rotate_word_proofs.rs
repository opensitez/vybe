use vybe_ast::{BinOp, BitLane, ExprKind, Expression, IntOp, Lang, Module, Statement, StmtKind};
use vybe_runtime::VM;
fn word(n: i64) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Integer(IntOp::Xor, BitLane::W32),
        left: Box::new(Expression::int(n)),
        right: Box::new(Expression::int(0)),
    })
}
fn compile(left: bool, a: Expression, b: Expression) -> Vec<vybe_runtime::Chunk> {
    let module = Module {
        name: "rotate-proof".into(),
        language: Lang::Unknown,
        body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
            ExprKind::Binary {
                op: if left {
                    BinOp::RotL(BitLane::W32)
                } else {
                    BinOp::RotR(BitLane::W32)
                },
                left: Box::new(a),
                right: Box::new(b),
            },
        ))))],
        imports: vec![],
        canon: Default::default(),
        directives: Default::default(),
    };
    vybe_compiler::Compiler::with_profile(
        vybe_compiler::profile::parse_profile("[compiler]\n").unwrap(),
    )
    .compile(&module)
    .unwrap()
}
#[test]
fn compiled_rotate_skips_only_proven_word_conversions() {
    for left in [false, true] {
        for a in [i32::MIN as i64, -1, 0, 1, i32::MAX as i64, u32::MAX as i64] {
            for b in [-65, -32, -1, 0, 1, 31, 32, 65] {
                for proof in [(false, false), (true, false), (false, true), (true, true)] {
                    let c = compile(
                        left,
                        if proof.0 { word(a) } else { Expression::int(a) },
                        if proof.1 { word(b) } else { Expression::int(b) },
                    );
                    let dump = vybe_runtime::debug::disassemble(&c[0]);
                    assert_eq!(
                        dump.matches("i64.trunc_f64_s").count(),
                        usize::from(!proof.0) + usize::from(!proof.1),
                        "{dump}"
                    );
                    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                        .validate_all(&vybe_platform_wasm::write_wasm(&c))
                        .unwrap();
                    let mut vm = VM::new();
                    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
                    let result = vm.run(c).unwrap();
                    let expected = if left {
                        (a as u32).rotate_left(b as u32)
                    } else {
                        (a as u32).rotate_right(b as u32)
                    };
                    assert_eq!(result.as_i32(), expected as i32);
                }
            }
        }
    }
}
#[test]
fn generic_fractional_and_invalid_numbers_keep_conversion() {
    for (a, b, expected) in [(4294967295.75, -1.75, -1), (1.75, 33.75, 2)] {
        let c = compile(true, Expression::float(a), Expression::float(b));
        let dump = vybe_runtime::debug::disassemble(&c[0]);
        assert_eq!(dump.matches("i64.trunc_f64_s").count(), 2);
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        assert_eq!(vm.run(c).unwrap().as_i32(), expected);
    }
    for n in [f64::NAN, f64::INFINITY] {
        for (a, b) in [
            (Expression::float(n), word(1)),
            (word(1), Expression::float(n)),
        ] {
            let mut vm = VM::new();
            vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
            assert!(vm.run(compile(true, a, b)).is_err());
        }
    }
}

#[test]
fn wide_rotate_raw_count_does_not_corrupt_later_boxed_scratch() {
    use vybe_runtime::{Chunk, Op};
    for left in [false, true] {
        let mut c = Chunk::new("wide-rotate-followed-by-value");
        c.emit_f64_const(-4294967296.0, 0);
        c.emit_f64_const(-1.0, 0);
        vybe_compiler::primitives::bits::emit_rotate(&mut c, BitLane::W64, left, 0);
        let answer = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, answer, 0);
        c.emit_f64_const(17.0, 0);
        let later = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, later, 0);
        c.emit_op_u16(Op::LOCAL_GET, answer, 0);
        c.emit_op(Op::RETURN, 0);
        assert_eq!(c.i64_locals.len(), 1);
        assert!(!c.i64_locals.contains(&answer));
        assert!(!c.i64_locals.contains(&later));
        let chunks = vec![c];
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        let expected = if left {
            (-4294967296i64 as u64).rotate_left(63)
        } else {
            (-4294967296i64 as u64).rotate_right(63)
        };
        assert_eq!(vm.run(chunks).unwrap().as_f64(), expected as i64 as f64);
    }
}

#[test]
fn wide_rotate_raw_slot_survives_statement_scratch_reclaim() {
    let rotate=Expression::new(ExprKind::Binary{op:BinOp::RotL(BitLane::W64),left:Box::new(Expression::float(1.0)),right:Box::new(Expression::float(63.0))});
    let module=Module{name:"wide-rotate-statement-reclaim".into(),language:Lang::Unknown,body:vec![Statement::new(StmtKind::Expr(rotate)),Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Binary{op:BinOp::RotR(BitLane::W32),left:Box::new(Expression::float(4.0)),right:Box::new(Expression::float(1.0))}))))],imports:vec![],canon:Default::default(),directives:Default::default()};
    let chunks=vybe_compiler::Compiler::with_profile(vybe_compiler::profile::parse_profile("[compiler]\n").unwrap()).compile(&module).unwrap();
    assert_eq!(chunks[0].i64_locals.len(),1);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&chunks)).unwrap();
    let mut vm=VM::new();vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);assert_eq!(vm.run(chunks).unwrap().as_i32(),2);
}
