use vybe_ast::{
    BinOp, BitLane, Directives, ExprKind, Expression, IntOp, Lang, Module, Statement, StmtKind,
    WideIntegerTransport,
};
use vybe_runtime::{Chunk, VM, Value};

fn cast(value: Expression, name: &str) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(value),
        type_name: name.into(),
    })
}

fn binary(operation: IntOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Integer(operation, BitLane::W64),
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn compile(value: Expression, transport: Option<WideIntegerTransport>) -> Vec<Chunk> {
    compile_with_bindings(value, transport, &[])
}

fn compile_with_bindings(
    value: Expression,
    transport: Option<WideIntegerTransport>,
    bindings: &[(&str, i64)],
) -> Vec<Chunk> {
    let profile = vybe_compiler::profile::parse_profile(
        "[compiler]\n[integer_cast_widths]\n\
         word = { bits = 64, signed = true }\n\
         uword = { bits = 64, signed = false }\n\
         small = { bits = 32, signed = true }\n\
         usmall = { bits = 32, signed = false }\n",
    )
    .unwrap();
    let mut body: Vec<_> = bindings
        .iter()
        .map(|(name, value)| {
            Statement::new(StmtKind::VarDecl {
                declarations: vec![vybe_ast::VarDeclarator {
                    pattern: vybe_ast::BindingPattern::Ident((*name).into()),
                    type_hint: Some("small".into()),
                    init: Some(Expression::int(*value)),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: vybe_ast::VarDeclKind::Var,
            })
        })
        .collect();
    body.push(Statement::new(StmtKind::Return(Some(value))));
    let module = Module {
        name: "exact-directive".into(),
        language: Lang::Unknown,
        body,
        imports: vec![],
        canon: Default::default(),
        directives: Directives {
            wide_integer_transport: transport,
            ..Default::default()
        },
    };
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}

fn run(chunks: Vec<Chunk>) -> Value {
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap()
}

#[test]
fn exact_wide_literals_and_arithmetic_keep_bits_above_number_precision() {
    for (value, expected) in [
        (
            cast(Expression::int(9_007_199_254_740_993), "word"),
            "9007199254740993",
        ),
        (cast(Expression::int(-1), "uword"), "18446744073709551615"),
        (
            binary(IntOp::Add, Expression::int(i64::MAX), Expression::int(1)),
            "-9223372036854775808",
        ),
        (
            binary(
                IntOp::Mul,
                Expression::int(9_007_199_254_740_993),
                Expression::int(3),
            ),
            "27021597764222979",
        ),
    ] {
        let chunks = compile(value, Some(WideIntegerTransport::Exact));
        assert!(
            chunks
                .iter()
                .all(|chunk| !vybe_runtime::debug::disassemble(chunk).contains("dyneq"))
        );
        let Value::BigInt(result) = run(chunks) else {
            panic!("missing exact carrier")
        };
        assert_eq!(result.to_string(), expected);
    }
}

#[test]
fn fixed_point_multiply_shift_and_small_cast_stay_native_without_bigint_bridge() {
    let left = 0x1234_5678i64;
    let right = -12345i64;
    let expression = cast(
        binary(
            IntOp::ShrS,
            binary(
                IntOp::Mul,
                cast(Expression::ident("left"), "word"),
                Expression::ident("right"),
            ),
            Expression::int(16),
        ),
        "small",
    );
    let chunks = compile_with_bindings(
        expression,
        Some(WideIntegerTransport::Exact),
        &[("left", left), ("right", right)],
    );
    let root = &chunks[0];
    assert!(
        !root
            .imports
            .iter()
            .any(|import| import.module == "ecma:bigint")
    );
    let dump = vybe_runtime::debug::disassemble(root);
    assert!(
        dump.contains("i64.mul") && dump.contains("i64.shr_s") && dump.contains("i32.wrap_i64"),
        "{dump}"
    );
    assert!(
        !dump.contains("dyneq") && !dump.contains("f64.convert_i64"),
        "{dump}"
    );
    assert_eq!(run(chunks), Value::I32(((left * right) >> 16) as i32));
}

#[test]
fn exact_small_casts_use_low_bits_without_whole_word_number_conversion() {
    let expression = cast(
        cast(Expression::int(0x1234_5678_ffff_ffff), "word"),
        "usmall",
    );
    assert_eq!(
        run(compile(expression, Some(WideIntegerTransport::Exact))).as_f64(),
        u32::MAX as f64
    );
}

#[test]
fn unstated_transport_keeps_the_existing_number_backed_path() {
    let expression = binary(IntOp::Add, Expression::int(1), Expression::int(2));
    let unstated = compile(expression.clone(), None);
    let explicit = compile(expression, Some(WideIntegerTransport::Number));
    assert_eq!(unstated[0].code, explicit[0].code);
    assert_eq!(run(unstated).as_f64(), 3.0);
}
