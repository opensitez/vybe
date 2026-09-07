use vybe_ast::{
    Argument, AtomicOp, AtomicRmw, BinOp, ExprKind, Expression, MemoryOrder, PlaceExpr, RmwResult,
    UnaryOp,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, methods, returns) in [
        (
            "atomic.Int32",
            &["Load", "Store", "Add", "Swap", "CompareAndSwap"][..],
            &[("Load", "int32"), ("Add", "int32"), ("Swap", "int32")][..],
        ),
        (
            "atomic.Int64",
            &["Load", "Store", "Add", "Swap", "CompareAndSwap"][..],
            &[("Load", "int64"), ("Add", "int64"), ("Swap", "int64")][..],
        ),
        (
            "atomic.Uint32",
            &["Load", "Store", "Add", "Swap", "CompareAndSwap"][..],
            &[("Load", "uint32"), ("Add", "uint32"), ("Swap", "uint32")][..],
        ),
        (
            "atomic.Uint64",
            &["Load", "Store", "Add", "Swap", "CompareAndSwap"][..],
            &[("Load", "uint64"), ("Add", "uint64"), ("Swap", "uint64")][..],
        ),
        (
            "atomic.Bool",
            &["Load", "Store", "Swap", "CompareAndSwap"][..],
            &[("Load", "bool"), ("Swap", "bool")][..],
        ),
        (
            "atomic.Value",
            &["Load", "Store"][..],
            &[("Load", "any")][..],
        ),
    ] {
        let mut method_tree = Subtree::new();
        for method in methods {
            method_tree.insert(
                (*method).to_string(),
                NamespaceNode::CommonEmit(format!("go.atomic.{name}.{method}")),
            );
        }
        insert_path(
            root,
            name,
            NamespaceNode::Type {
                ctor: None,
                ctor_call: None,
                statics: Subtree::new(),
                methods: method_tree,
                member_returns: returns
                    .iter()
                    .map(|(member, ty)| ((*member).to_string(), (*ty).to_string()))
                    .collect(),
            },
        );
    }
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "atomic.Int32" | "sync/atomic.Int32" | "sync.atomic.Int32" => Some("__goAtomicInt32"),
        "atomic.Int64" | "sync/atomic.Int64" | "sync.atomic.Int64" => Some("__goAtomicInt64"),
        "atomic.Uint32" | "sync/atomic.Uint32" | "sync.atomic.Uint32" => Some("__goAtomicUint32"),
        "atomic.Uint64" | "sync/atomic.Uint64" | "sync.atomic.Uint64" => Some("__goAtomicUint64"),
        "atomic.Bool" | "sync/atomic.Bool" | "sync.atomic.Bool" => Some("__goAtomicBool"),
        "atomic.Value" | "sync/atomic.Value" | "sync.atomic.Value" => Some("__goAtomicValue"),
        _ => None,
    }
}

/// Go `sync/atomic` function-style operations normalized to the shared atomic
/// primitive. Go's package surface owns the spellings; the concurrency model is
/// common AST.
pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let rest = call_name.strip_prefix("atomic.")?;
    let ordering = MemoryOrder::SeqCst;
    let place = || atomic_place(arg_value(args, 0)).map(Box::new);
    let atomic = if rest.starts_with("Load") && args.len() == 1 {
        AtomicOp::Load {
            place: place()?,
            ordering,
        }
    } else if rest.starts_with("Store") && args.len() == 2 {
        AtomicOp::Store {
            place: place()?,
            value: Box::new(arg_value(args, 1)),
            ordering,
        }
    } else if rest.starts_with("Add") && args.len() == 2 {
        AtomicOp::Rmw {
            op: AtomicRmw::Add,
            place: place()?,
            operand: Box::new(arg_value(args, 1)),
            result: RmwResult::New,
            ordering,
        }
    } else if rest.starts_with("Swap") && args.len() == 2 {
        AtomicOp::Rmw {
            op: AtomicRmw::Xchg,
            place: place()?,
            operand: Box::new(arg_value(args, 1)),
            result: RmwResult::Old,
            ordering,
        }
    } else if rest.starts_with("CompareAndSwap") && args.len() == 3 {
        let expected = arg_value(args, 1);
        let replacement = arg_value(args, 2);
        return Some(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::new(ExprKind::Atomic(
                AtomicOp::CompareExchange {
                    place: place()?,
                    expected: Box::new(expected.clone()),
                    replacement: Box::new(replacement),
                    result: RmwResult::Old,
                    ordering,
                },
            ))),
            op: BinOp::Eq,
            right: Box::new(expected),
        }));
    } else {
        return None;
    };
    Some(Expression::new(ExprKind::Atomic(atomic)))
}

pub(crate) fn rewrite_typed_method(
    receiver: Expression,
    receiver_type: &str,
    method: &str,
    args: &[Argument],
) -> Option<Expression> {
    let ty = receiver_type
        .trim()
        .trim_start_matches('*')
        .trim_start_matches('^')
        .trim();
    let value = Expression::new(ExprKind::Member {
        object: Box::new(receiver),
        field: "value".to_string(),
        null_safe: false,
    });

    if ty == "__goAtomicValue" {
        return match (method, args.len()) {
            ("Load", 0) => Some(value),
            ("Store", 1) => Some(Expression::new(ExprKind::Assign {
                target: Box::new(value),
                value: Box::new(arg_value(args, 0)),
            })),
            _ => None,
        };
    }

    let suffix = match ty {
        "__goAtomicInt32" => "Int32",
        "__goAtomicInt64" => "Int64",
        "__goAtomicUint32" => "Uint32",
        "__goAtomicUint64" => "Uint64",
        "__goAtomicBool" => "Bool",
        _ => return None,
    };
    let call_name = format!("atomic.{method}{suffix}");
    let mut lowered_args = Vec::with_capacity(args.len() + 1);
    lowered_args.push(Argument::positional(value));
    lowered_args.extend(args.iter().cloned());
    rewrite_call(&call_name, &lowered_args)
}

fn arg_value(args: &[Argument], index: usize) -> Expression {
    args.get(index)
        .map(|arg| arg.value.clone())
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(vybe_ast::Literal::Null)))
}

fn atomic_place(arg: Expression) -> Option<Expression> {
    match arg.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => Some(*expr),
        ExprKind::RefOf(place) => Some(place_expr(&place)),
        ExprKind::Ident(_) => Some(arg),
        _ => PlaceExpr::from_expr(&arg).map(|place| place_expr(&place)),
    }
}

fn place_expr(place: &PlaceExpr) -> Expression {
    match place {
        PlaceExpr::Ident(name) => Expression::ident(name),
        PlaceExpr::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: object.clone(),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: object.clone(),
            index: index.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Deref(expr) => Expression::new(ExprKind::RefLoad(expr.clone())),
    }
}

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let Some(leaf) = segments.pop() else {
        return;
    };
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.insert(leaf.to_string(), node);
}
