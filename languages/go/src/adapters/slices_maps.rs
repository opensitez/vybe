use vybe_ast::{Argument, ArrayElement, BinOp, ExprKind, Expression};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("slices.Contains", "collections.contains"),
        ("slices.Index", "collections.index_of"),
        ("slices.IndexFunc", "collections.index_func"),
        ("slices.Equal", "collections.sequence_equal"),
        ("slices.Compare", "collections.sequence_compare"),
        ("slices.Clone", "collections.clone"),
        ("slices.Compact", "collections.compact_adjacent"),
        ("slices.Delete", "collections.delete_range_copy"),
        ("slices.DeleteFunc", "go.slices_delete_func"),
        ("slices.Grow", "collections.first_of_two"),
        ("slices.Clip", "collections.clone"),
        ("slices.CompactFunc", "go.slices_compact_func"),
        ("slices.Values", "go.slices_values"),
        ("slices.Sort", "collections.sort"),
        ("slices.SortFunc", "collections.sort_func"),
        ("slices.SortStableFunc", "collections.sort_func"),
        ("slices.IsSorted", "collections.is_sorted"),
        ("slices.IsSortedFunc", "collections.is_sorted_func"),
        ("slices.BinarySearch", "collections.binary_search_pair"),
        (
            "slices.BinarySearchFunc",
            "collections.binary_search_func_pair",
        ),
        ("maps.Clone", "collections.map_clone"),
        ("maps.Copy", "collections.map_copy"),
        ("maps.DeleteFunc", "collections.map_delete_func"),
        ("maps.Keys", "go.maps_keys"),
        ("maps.Values", "go.maps_values"),
        ("maps.Equal", "go.maps_equal"),
        ("maps.EqualFunc", "go.maps_equal_func"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }
}

/// Go `slices`/`maps` package calls normalized to the shared collection
/// primitive surface where it already exists, with Go-owned adapter leaves for
/// package behavior that needs a thin semantic wrapper.
pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let direct = |helper: &str, args: &[Argument]| {
        Some(go_builtin_call(
            helper,
            args.iter().map(|a| a.value.clone()).collect(),
        ))
    };
    let direct_with_callable = |helper: &str, args: &[Argument], callable_idx: usize| {
        Some(go_builtin_call(
            helper,
            args.iter()
                .enumerate()
                .map(|(idx, a)| {
                    if idx == callable_idx {
                        unwrap_callable_cast(a.value.clone())
                    } else {
                        a.value.clone()
                    }
                })
                .collect(),
        ))
    };
    let direct_nil_empty = |helper: &str, args: &[Argument]| {
        Some(go_builtin_call(
            helper,
            args.iter()
                .map(|a| go_slice_nil_to_empty_expr(a.value.clone()))
                .collect(),
        ))
    };
    match call_name {
        "slices.Contains" => direct("__go_slices_contains_common", args),
        "slices.Index" => direct("__go_slices_index_common", args),
        "slices.IndexFunc" => direct_with_callable("__go_slices_index_func_common", args, 1),
        "slices.Equal" => direct_nil_empty("__go_slices_equal_common", args),
        "slices.Compare" => direct_nil_empty("__go_slices_compare_common", args),
        "slices.Clone" => args.first().map(|arg| {
            let slice = arg.value.clone();
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(slice.clone()),
                    right: Box::new(Expression::null()),
                })),
                then: Box::new(Expression::null()),
                else_: Box::new(go_builtin_call("__go_slices_clone_common", vec![slice])),
            })
        }),
        "slices.Compact" => direct("__go_slices_compact_common", args),
        "slices.Delete" => direct("__go_slices_delete_common", args),
        "slices.DeleteFunc" => direct_with_callable("go.slices_delete_func", args, 1),
        "slices.CompactFunc" => direct_with_callable("go.slices_compact_func", args, 1),
        "slices.Grow" => args.first().map(|_| {
            go_builtin_call(
                "__go_slices_grow_common",
                args.iter().take(2).map(|a| a.value.clone()).collect(),
            )
        }),
        "slices.Clip" => direct("__go_slices_clip_common", args),
        "slices.All" => args.first().map(|a| a.value.clone()),
        "slices.Values" => direct("go.slices_values", args),
        "slices.Sort" => direct("__go_slices_sort_common", args),
        "slices.SortFunc" => direct_with_callable("__go_slices_sort_func_common", args, 1),
        "slices.SortStableFunc" => direct_with_callable("__go_slices_sort_func_common", args, 1),
        "slices.IsSorted" => direct("__go_slices_is_sorted_common", args),
        "slices.IsSortedFunc" => direct_with_callable("__go_slices_is_sorted_func_common", args, 1),
        "slices.BinarySearch" => {
            let mut call_args: Vec<Expression> = args.iter().map(|a| a.value.clone()).collect();
            if let Some(first) = call_args.first_mut() {
                *first = go_slice_nil_to_empty_expr(first.clone());
            }
            Some(go_builtin_call(
                "__go_slices_binary_search_pair_common",
                call_args,
            ))
        }
        "slices.BinarySearchFunc" => {
            let mut call_args: Vec<Expression> = args.iter().map(|a| a.value.clone()).collect();
            if let Some(first) = call_args.first_mut() {
                *first = go_slice_nil_to_empty_expr(first.clone());
            }
            if let Some(arg) = call_args.get_mut(2) {
                *arg = unwrap_callable_cast(arg.clone());
            }
            Some(go_builtin_call(
                "__go_slices_binary_search_func_pair_common",
                call_args,
            ))
        }
        "maps.Equal" => direct("go.maps_equal", args),
        "maps.EqualFunc" => direct_with_callable("go.maps_equal_func", args, 2),
        "maps.Clone" => direct("__go_maps_clone_common", args),
        "maps.Copy" => direct("__go_maps_copy_common", args),
        "maps.DeleteFunc" => direct("__go_maps_delete_func_common", args),
        "slices.Insert" => {
            let head: Vec<Expression> = args.iter().take(2).map(|a| a.value.clone()).collect();
            let mut call_args = head;
            call_args.push(go_variadic_tail_as_slice(args, 2));
            Some(go_builtin_call("__go_slices_insert_common", call_args))
        }
        "slices.Replace" => {
            let head: Vec<Expression> = args.iter().take(3).map(|a| a.value.clone()).collect();
            let mut call_args = head;
            call_args.push(go_variadic_tail_as_slice(args, 3));
            Some(go_builtin_call("__go_slices_replace_common", call_args))
        }
        _ => None,
    }
}

fn unwrap_callable_cast(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::Cast { expr, .. } => *expr,
        _ => expr,
    }
}

fn go_builtin_call(name: &str, args: Vec<Expression>) -> Expression {
    let name = private_helper_name(name);
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn private_helper_name(name: &str) -> &str {
    match name {
        "go.slices_delete_func" => "__go_slices_delete_func",
        "go.slices_compact_func" => "__go_slices_compact_func",
        "go.maps_equal" => "__go_maps_equal",
        "go.maps_equal_func" => "__go_maps_equal_func",
        _ => name,
    }
}

fn go_array_of(elems: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        elems
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn go_slice_nil_to_empty_expr(expr: Expression) -> Expression {
    go_builtin_call("__go_slices_nil_to_empty_common", vec![expr])
}

fn go_variadic_tail_as_slice(args: &[Argument], start: usize) -> Expression {
    if args.len() == start + 1 && args[start].spread {
        args[start].value.clone()
    } else {
        go_array_of(args.iter().skip(start).map(|a| a.value.clone()).collect())
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
