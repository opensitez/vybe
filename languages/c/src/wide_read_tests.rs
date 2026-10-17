use vybe_ast::{BinOp, BitLane, ExprKind, IntOp, StmtKind};

#[test]
fn wide_scalar_reads_preserve_comparison_lanes_without_reclassifying_floats() {
    for (source, expected) in [
        ("int f(long long *p){return p[1]!=50000000005LL;}", Some(IntOp::Ne)),
        ("int f(long long *p){return *p < -10000000001LL;}", Some(IntOp::LtS)),
        ("int f(unsigned long long *p){return p[1]>0ULL;}", Some(IntOp::GtU)),
        ("typedef unsigned long long W;int f(W *p){return *p!=1ULL;}", Some(IntOp::Ne)),
        ("int f(double *p){return p[1]!=50000000005LL;}", None),
    ] {
        for parse in [crate::parse_with_peg, crate::parse_optimized] {
            let module = parse(source).unwrap();
            let body = module.body.iter().find_map(|statement| match &statement.kind {
                StmtKind::FunctionDecl { name, body, .. } if name == "f" => Some(body),
                _ => None,
            }).unwrap();
            let mut comparisons = Vec::new();
            for statement in body {
                statement.walk_exprs(&mut |value| {
                    if let ExprKind::Binary { op: BinOp::Integer(op, BitLane::W64), .. } = &value.kind
                        && op.is_comparison()
                    {
                        comparisons.push(*op);
                    }
                });
            }
            if let Some(expected) = expected {
                assert!(comparisons.contains(&expected), "{source}: {comparisons:?}");
            } else {
                assert!(comparisons.is_empty(), "float read acquired an integer comparison: {source}");
            }
        }
    }
}
