use std::collections::{HashMap, HashSet};
use vybe_ast::{BinOp, BitLane, ExprKind, IntOp, StmtKind};

#[test]
fn source_pointer_ordering_keeps_typed_coordinates_through_pooled_helpers() {
    let module = vybe_language_c::parse(include_str!(
        "../../../tests/c/doom_startup/static_pointer_ordering.c"
    ))
    .unwrap();
    let functions: HashMap<_, _> = module
        .body
        .iter()
        .filter_map(|statement| {
            if let StmtKind::FunctionDecl { name, .. } = &statement.kind {
                Some((name.as_str(), statement))
            } else {
                None
            }
        })
        .collect();
    for name in ["after", "ordered", "record_after", "evaluated_once"] {
        let mut managed = 0;
        let mut numeric = 0;
        let mut calls = 0;
        // parse() can hoist IIFEs; inspect their actual bodies, once per helper.
        let mut pending = vec![name.to_string()];
        let mut visited = HashSet::new();
        while let Some(function_name) = pending.pop() {
            if !visited.insert(function_name.clone()) {
                continue;
            }
            let function = functions
                .get(function_name.as_str())
                .unwrap_or_else(|| panic!("missing AST function {function_name}"));
            function.walk_exprs(&mut |value| {
                if let ExprKind::Call { callee, args, .. } = &value.kind
                    && let ExprKind::Ident(callee) = &callee.kind
                {
                    let target = if callee == "__c_invoke_fixed" {
                        args.first()
                            .and_then(|argument| match &argument.value.kind {
                                ExprKind::Ident(name) => Some(name.as_str()),
                                _ => None,
                            })
                    } else {
                        Some(callee.as_str())
                    };
                    if let Some(target) = target {
                        calls += usize::from(target == "position");
                        if target.starts_with("__c_static_iife_") {
                            pending.push(target.to_string());
                        }
                    }
                }
                if let ExprKind::Binary { op, left, right } = &value.kind {
                    let mut coordinate = false;
                    for operand in [left, right] {
                        operand.walk_exprs(&mut |value| {
                            coordinate |= matches!(&value.kind,
                            ExprKind::Member { field, .. } if field == "__idx");
                        });
                    }
                    if matches!(op, BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq) {
                        assert!(!coordinate, "untyped managed ordering in {name}");
                    }
                    if matches!(
                        op,
                        BinOp::Integer(
                            IntOp::LtS | IntOp::LeS | IntOp::GtS | IntOp::GeS,
                            BitLane::W64
                        )
                    ) && coordinate
                    {
                        managed += 1;
                    }
                    if matches!(
                        op,
                        BinOp::Integer(
                            IntOp::LtU | IntOp::LeU | IntOp::GtU | IntOp::GeU,
                            BitLane::W32
                        )
                    ) {
                        numeric += 1;
                    }
                }
            });
        }
        assert!(managed > 0, "missing signed managed coordinates in {name}");
        if name == "after" {
            assert!(numeric > 0, "missing numeric address arm");
        }
        assert_eq!(
            calls,
            if name == "evaluated_once" { 2 } else { 0 },
            "pointer operands must not be duplicated in {name}"
        );
    }
}
