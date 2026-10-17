#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn nested_float_arithmetic_retains_its_result_lane() {
    use vybe_ast::{BinOp, ExprKind, StmtKind};
    for source in NUMERIC_PROGRAMS {
        for parse in [crate::parse_with_peg, crate::parse_optimized] {
            let module = parse(source).unwrap();
            let main = module
                .body
                .iter()
                .find(|statement| {
                    matches!(&statement.kind,
            StmtKind::FunctionDecl { name, .. } if name == "main")
                })
                .unwrap();
            main.walk_exprs(&mut |value| {
                assert!(
                    !matches!(
                        &value.kind,
                        ExprKind::Binary {
                            op: BinOp::IDiv
                                | BinOp::Integer(vybe_ast::IntOp::DivS | vybe_ast::IntOp::DivU, _),
                            ..
                        }
                    ),
                    "floating ratio became integer division: {value:?}"
                );
            });
        }
    }
}

const NUMERIC_PROGRAMS: &[&str] = &[
    "int main(void){volatile int ticks=4,elapsed=144; \
        float ratio=((float)ticks*35)/elapsed;return ratio<0.97f||ratio>0.98f;}",
    "double ticks(void){return 4.0;} int main(void){volatile int elapsed=144; \
        double ratio=ticks()*35/elapsed;return ratio<0.97||ratio>0.98;}",
    "struct Counter {double ticks;}; int main(void){struct Counter c={4.0}; \
        double a[1]={35.0};double ratio=c.ticks*a[0]/144; \
        return ratio<0.97||ratio>0.98;}",
    "int main(void){volatile long long scale=35;volatile double ticks=4.0; \
        double ratio=ticks*scale/144;return ratio<0.97||ratio>0.98;}",
];

#[test]
fn declared_float_operands_execute_without_integer_division_or_dynamic_equality() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(
        crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        NUMERIC_PROGRAMS.iter().copied(),
    );
}

#[test]
fn float_ratio_and_default_variadic_promotion_execute_standard_wasm_without_dynamic_equality() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(
        crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        [include_str!(
            "../../../tests/c/doom_startup/float_ratio_varargs.c"
        )],
    );
}
