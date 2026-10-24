use super::*;

const OVERFLOW_FIXTURE: &str =
    include_str!("../../../tests/c/c_calloc_overflow/calloc_overflow.c");

fn initializer<'a>(module: &'a Module, function: &str, binding: &str) -> &'a Expression {
    let body = module.body.iter().find_map(|statement| match &statement.kind {
        StmtKind::FunctionDecl { name, body, .. } if name == function => Some(body),
        _ => None,
    }).expect("function body");
    body.iter().find_map(|statement| match &statement.kind {
        StmtKind::VarDecl { declarations, .. } => declarations.iter().find_map(|declaration| {
            matches!(&declaration.pattern, BindingPattern::Ident(name) if name == binding)
                .then_some(declaration.init.as_ref()).flatten()
        }),
        _ => None,
    }).expect("pointer initializer")
}

fn call_count(value: &Expression, name: &str) -> usize {
    let statement = stmt(StmtKind::Expr(value.clone()));
    let mut count = 0;
    statement.walk_exprs(&mut |value| {
        if matches!(&value.kind, ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(callee) if callee == name)) {
            count += 1;
        }
    });
    count
}

fn assert_saved_fixture_has_no_allocation() {
    for parse in [crate::parse_with_peg, crate::parse_optimized] {
        let module = parse(OVERFLOW_FIXTURE).unwrap();
        let value = initializer(&module, "main", "p");
        assert!(matches!(value.kind, ExprKind::Lit(Literal::Null)), "{value:?}");
        for allocator in ["Array", "__c_ptr_linear_alloc", "__c_u8array_new"] {
            assert_eq!(call_count(value, allocator), 0, "{allocator}: {value:?}");
        }
    }
}

#[test]
fn exact_saved_overflow_fixture_initializes_null_not_an_array_allocator() {
    assert_saved_fixture_has_no_allocation();
}

#[test]
fn constant_rejection_uses_storage_bound_not_signed_product_overflow() {
    let walker = Walker::default();
    // These products fit size_t/u64 but not the wasm32 storage representation.
    for (count, size) in [(i64::MAX, 2), (65536, 65536), (-1, 1)] {
        assert!(walker.reject_constant_calloc_request(&int_lit(count), &int_lit(size)));
    }
    for (count, size) in [(0, 8), (8, 0), (2, 4), (64000, 1), (1000, 1000)] {
        assert!(!walker.reject_constant_calloc_request(&int_lit(count), &int_lit(size)));
    }
    assert!(!walker.reject_constant_calloc_request(&int_lit(i64::MAX), &int_lit(0)));
    for value in [f64::INFINITY, f64::NEG_INFINITY, f64::NAN] {
        assert!(walker.reject_constant_calloc_request(&float_lit(value), &int_lit(1)));
    }
}

#[test]
fn integer_cancellation_is_not_approximated_as_floating_point() {
    let walker = Walker::default();
    let count = binary_expr(BinOp::Sub,
        binary_expr(BinOp::Add, int_lit(9007199254740993), int_lit(1)),
        int_lit(9007199254740993));
    assert_eq!(walker.calloc_literal_number(&count), Some(1.0));
    assert!(!walker.reject_constant_calloc_request(&count, &int_lit(10000000)));
}

#[test]
fn source_constant_proof_precedes_fenv_but_never_accepts_calls_or_casts() {
    let walker = Walker::default();
    let source = "9223372036854775807.0 / 2 + 2";
    let mut pairs = CParser::parse(Rule::assignment_expression, source).unwrap();
    let literal = walker.calloc_source_literal(pairs.next().unwrap()).unwrap();
    assert!(walker.reject_constant_calloc_request(&literal, &int_lit(2)));
    for source in ["count_once()", "n++", "(size_t)n", "(n=3)", "(n,3)",
        "__c_fenv_binary(0,9223372036854775807.0,2)"] {
        let mut pairs = CParser::parse(Rule::assignment_expression, source).unwrap();
        assert!(walker.calloc_source_literal(pairs.next().unwrap()).is_none(), "{source}");
    }
    let helper = expr(ExprKind::Cast {
        expr: Box::new(call_expr(ident("__c_fenv_binary"),
            vec![int_lit(0), float_lit(i64::MAX as f64), int_lit(2)])),
        type_name: "double".into(),
    });
    assert!(!walker.reject_constant_calloc_request(&helper, &int_lit(2)));
}

#[test]
fn unknown_and_effectful_arguments_are_not_constant_rejection_proofs() {
    let walker = Walker::default();
    let effect = call_expr(ident("count_once"), Vec::new());
    assert!(!walker.reject_constant_calloc_request(&effect, &int_lit(i64::MAX)));
    assert!(!walker.reject_constant_calloc_request(&ident("n"), &ident("size")));
    for parse in [crate::parse_with_peg, crate::parse_optimized] {
        let module = parse(
            "#include <stdlib.h>\nint count_once(void){return 2;} \
             int size_once(void){return 4;} \
             int main(void){void *p=calloc(count_once(),size_once());return p==0;}",
        ).unwrap();
        let value = initializer(&module, "main", "p");
        assert_eq!(call_count(value, "count_once"), 1, "{value:?}");
        assert_eq!(call_count(value, "size_once"), 1, "{value:?}");
    }
}

#[test]
fn source_constant_rejection_does_not_override_a_user_calloc_binding() {
    for parse in [crate::parse_with_peg, crate::parse_optimized] {
        let module = parse(
            "void *calloc(double n,int size){return 0;} \
             int main(void){void *p=calloc(9223372036854775807.0/2+2,2);return p!=0;}",
        ).unwrap();
        assert_eq!(call_count(initializer(&module, "main", "p"), "calloc"), 1);
    }
}

#[test]
#[ignore = "dynamic calloc guard not implemented; AST-only desired-behavior regression"]
fn dynamic_overflow_requires_a_null_failure_branch_and_single_argument_evaluation() {
    // Parse only. Never execute these overflow requests in the Rust test VM.
    // The missing runtime guard must eventually make NULL reachable without
    // evaluating the allocator, while evaluating each C argument exactly once.
    for parse in [crate::parse_with_peg, crate::parse_optimized] {
        let module = parse(
            "#include <stdlib.h>\nsize_t count_once(void){return (size_t)-1;} \
             size_t size_once(void){return 2;} \
             int main(void){void *p=calloc(count_once(),size_once());return p!=NULL;}",
        ).unwrap();
        let value = initializer(&module, "main", "p");
        assert_eq!(call_count(value, "count_once"), 1);
        assert_eq!(call_count(value, "size_once"), 1);
        let mut null_failure = false;
        stmt(StmtKind::Expr(value.clone())).walk_exprs(&mut |value| {
            if let ExprKind::Ternary { then, else_, .. } = &value.kind {
                null_failure |= matches!(then.kind, ExprKind::Lit(Literal::Null))
                    || matches!(else_.kind, ExprKind::Lit(Literal::Null));
            }
        });
        assert!(null_failure, "dynamic calloc needs an allocation-skipping NULL branch: {value:?}");
    }
}

const VALID_SHAPES: &str = "#include <stdlib.h>\n\
    struct Pair{int first;int second;}; \
    int main(void){int *p=calloc(2,sizeof(int)); \
    unsigned char *b=calloc(4,1); struct Pair *s=calloc(1,sizeof(struct Pair)); \
    p[1]=7;b[3]=9;s->second=11; \
    return p[0]!=0 || p[1]!=7 || b[0]!=0 || b[3]!=9 || s->first!=0 || s->second!=11;}";

const EVALUATE_ONCE: &str = "#include <stdlib.h>\n\
    static int counts;static int sizes; \
    int count_once(void){counts++;return 2;} \
    int size_once(void){sizes++;return 4;} \
    int main(void){int *p=calloc(count_once(),size_once());p[1]=7; \
    return counts!=1 || sizes!=1 || p[0]!=0 || p[1]!=7;}";

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
#[ignore = "run AST-only exact-fixture gate first, then coordinate bounded external execution"]
fn constant_calloc_and_valid_shapes_execute_standard_wasm_without_dynamic_equality() {
    // Independent precondition: a separately failing AST test cannot protect
    // this process if a test runner selects execution tests directly.
    assert_saved_fixture_has_no_allocation();
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        [OVERFLOW_FIXTURE, VALID_SHAPES, EVALUATE_ONCE]);
}
