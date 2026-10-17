use vybe_ast::*;
use vybe_compiler::primitives::Compiler;
use vybe_runtime::{Chunk, VM, Value};

const PROFILES: [&str; 2] = [
    include_str!("../../../languages/go/src/profile"),
    include_str!("../../../languages/php/src/profile"),
];

fn function(name: &str, parameters: bool, body: Vec<Statement>) -> Statement {
    Statement::new(StmtKind::FunctionDecl {
        name: name.into(),
        params: if parameters {
            vec![Param {
                name: "value".into(),
                type_hint: Some("int32".into()),
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }]
        } else {
            Vec::new()
        },
        return_type: None,
        body,
        modifiers: Default::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}

fn call(name: &str, arguments: bool) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: if arguments {
            vec![Argument::positional(Expression::int(3))]
        } else {
            Vec::new()
        },
        optional: false,
    })
}

fn module(target: &str, called: &str, guest_body: Vec<Statement>, immutable: bool) -> Module {
    let guest = Expression::new(ExprKind::FunctionExpr(Box::new(function(
        "guest", false, guest_body,
    ))));
    Module {
        name: "expression-publications".into(),
        language: Lang::Unknown,
        imports: Vec::new(),
        canon: Default::default(),
        directives: Directives {
            declared_function_bindings_live: Some(!immutable),
            callable_invocation: Some(CallableInvocation::Single),
            receiver_binding: Some(ReceiverBinding::UniversalParameter),
            ..Default::default()
        },
        body: vec![
            function(
                target,
                true,
                vec![Statement::new(StmtKind::Return(Some(Expression::ident(
                    "value",
                ))))],
            ),
            function(
                "caller",
                false,
                vec![Statement::new(StmtKind::Return(Some(call(called, true))))],
            ),
            function(
                "startup",
                false,
                vec![Statement::new(StmtKind::Expr(guest))],
            ),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident("publication_result")],
                value: call("caller", false),
                by_ref: false,
            }),
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(Expression::ident("publication_result")),
                    right: Box::new(Expression::int(3)),
                }),
                then_body: vec![Statement::new(StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::Lit(Literal::Str(
                        "wrong publication result".into(),
                    )))),
                    cause: None,
                })],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Return(Some(Expression::ident(
                "publication_result",
            )))),
        ],
    }
}

fn compile(profile: vybe_runtime::profile::LanguageProfile, module: &Module) -> Vec<Chunk> {
    Compiler::with_profile(profile)
        .with_run_module_end_hooks(false)
        .compile(module)
        .unwrap()
}

fn numeric_three(value: Value) {
    assert!(
        value == Value::I32(3) || value == Value::F64(3.0),
        "unexpected result: {value:?}"
    );
}

fn vm() -> VM {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    vm
}

fn standard_wasm(chunks: &[Chunk]) -> Vec<u8> {
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    bytes
}

fn verify_standard_wasm_and_execution(chunks: &[Chunk]) {
    numeric_three(vm().run(chunks.to_vec()).unwrap());
    let bytes = standard_wasm(chunks);
    let decoded = vybe_platform_wasm::read_wasm(&bytes).unwrap();
    let mut reloaded = vm();
    reloaded.run(decoded).unwrap();
    assert!(reloaded.last_exception.is_none());
    // The reader's _start adapter discards results. The source module throws
    // unless the caller returned three; a separate wrong-result control below
    // proves the reloaded path does not bypass that check.
}

#[test]
fn unrelated_function_expression_retains_real_completed_edge_and_standard_execution() {
    for text in PROFILES {
        let profile = vybe_compiler::profile::parse_profile(text).unwrap();
        for immutable in [false, true] {
            let source = module(
                "target",
                "target",
                vec![Statement::new(StmtKind::Return(None))],
                immutable,
            );
            let chunks = compile(profile.clone(), &source);
            let target = chunks
                .iter()
                .position(|chunk| chunk.name == "target")
                .unwrap();
            let caller = chunks.iter().find(|chunk| chunk.name == "caller").unwrap();
            if immutable {
                assert_eq!(
                    caller.direct_defined_calls.len(),
                    1,
                    "{} must admit the actual target entry",
                    profile.name
                );
                let edge = caller.direct_defined_calls.values().next().unwrap();
                assert_eq!(edge.target as usize, target);
                assert!(edge.source);
                assert_eq!(chunks[target].direct_entry_has_captures, Some(false));
            } else {
                assert!(caller.direct_defined_calls.is_empty());
            }
            verify_standard_wasm_and_execution(&chunks);
        }
    }
}

#[test]
fn nested_conditional_alias_collision_retains_callable_path_and_standard_execution() {
    for text in PROFILES {
        let mut profile = vybe_compiler::profile::parse_profile(text).unwrap();
        // Exercise the actual shared alias publisher for both profiles.
        profile.source_function_callable_aliases = true;
        let conditional = Statement::new(StmtKind::If {
            cond: Expression::bool(false),
            then_body: vec![function("Outer.Inner", false, vec![])],
            elifs: vec![],
            else_body: None,
        });
        for immutable in [false, true] {
            let source = module(
                "__php_fn_Outer__Inner",
                "__php_fn_Outer__Inner",
                vec![conditional.clone()],
                immutable,
            );
            let chunks = compile(profile.clone(), &source);
            let caller = chunks.iter().find(|chunk| chunk.name == "caller").unwrap();
            assert!(
                caller.direct_defined_calls.is_empty(),
                "{} must not certify a colliding alias",
                profile.name
            );
            verify_standard_wasm_and_execution(&chunks);
        }
    }
}

#[test]
fn metadata_free_reload_really_executes_the_publication_result_check() {
    for text in PROFILES {
        let profile = vybe_compiler::profile::parse_profile(text).unwrap();
        let mut source = module("target", "target", vec![], false);
        let check = source
            .body
            .iter_mut()
            .find(|statement| matches!(statement.kind, StmtKind::If { .. }))
            .unwrap();
        let StmtKind::If { cond, .. } = &mut check.kind else {
            unreachable!()
        };
        let ExprKind::Binary { right, .. } = &mut cond.kind else {
            unreachable!()
        };
        **right = Expression::int(4);
        let chunks = compile(profile, &source);
        let mut ordinary = vm();
        assert!(ordinary.run(chunks.clone()).is_err() || ordinary.last_exception.is_some());
        let decoded = vybe_platform_wasm::read_wasm(&standard_wasm(&chunks)).unwrap();
        let mut reloaded = vm();
        assert!(reloaded.run(decoded).is_err() || reloaded.last_exception.is_some());
    }
}
