#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_runtime::{VM, Value};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;

fn binary(op: IntOp, lane: BitLane, a: Expression, b: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Integer(op, lane),
        left: Box::new(a),
        right: Box::new(b),
    })
}

#[test]
fn w64_sum_is_wrapped_only_by_its_w32_consumer() {
    let points = [i32::MIN, -1, 0, 1, i32::MAX];
    let mut checks = 0;
    for (oi, op) in [IntOp::Add, IntOp::Sub].into_iter().enumerate() {
        for (ai, a) in points.into_iter().enumerate() {
            for (bi, b) in points.into_iter().enumerate() {
                let sum = if op == IntOp::Add {
                    i64::from(a) + i64::from(b)
                } else {
                    i64::from(a) - i64::from(b)
                };
                for fast in [false, true] {
                    let inner = binary(
                        op,
                        BitLane::W64,
                        Expression::int(i64::from(a)),
                        Expression::int(i64::from(b)),
                    );
                    let expression = binary(IntOp::Or, BitLane::W32, inner, Expression::int(0));
                    let module = Module {
                        name: "sum-consumer".into(),
                        language: Lang::Unknown,
                        imports: vec![],
                        canon: Default::default(),
                        directives: Default::default(),
                        body: vec![Statement::new(StmtKind::Return(Some(expression)))],
                    };
                    let profile = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
                    let chunks = vybe_compiler::primitives::bits::benchmark_with_wide_signed_words(
                        fast,
                        || {
                            vybe_compiler::Compiler::with_profile(profile)
                                .compile(&module)
                                .unwrap()
                        },
                    );
                    let bytes = vybe_platform_wasm::write_wasm_with_options(
                        &chunks,
                        vybe_platform_wasm::WasmWriteOptions {
                            include_name_section: false,
                            include_vybe_metadata: false,
                        },
                    );
                    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                        .validate_all(&bytes)
                        .unwrap();
                    native_export::export_chunks(
                        &chunks,
                        &format!("sum-consumer-{oi}-{ai}-{bi}-{fast}"),
                    );
                    assert_eq!(
                        VM::new().run(chunks).unwrap(),
                        Value::I32(sum as i32),
                        "op={op:?}, a={a}, b={b}, fast={fast}"
                    );
                    checks += 1;
                }
            }
        }
    }
    let directory = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::write(
        directory.join("W64_SUMS_CONSUMER_GATES.json"),
        serde_json::to_vec_pretty(
            &serde_json::json!({"checks":checks,"native_validations":checks}),
        )
        .unwrap(),
    )
    .unwrap();
}
