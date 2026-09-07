use vybe_ast::{Argument, ArrayElement, ExprKind, Expression};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("errors.New", "go.errors_new"),
        ("errors.Unwrap", "go.errors_unwrap"),
        ("errors.Is", "go.errors_is"),
        ("errors.Join", "go.errors_join"),
        ("errors.As", "go.errors_as"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "errors.New" => Some(go_builtin_call("__go_errors_new", vec![arg_value(args, 0)])),
        "errors.Unwrap" => Some(go_builtin_call(
            "__go_errors_unwrap",
            vec![arg_value(args, 0)],
        )),
        "errors.Is" => Some(go_builtin_call(
            "__go_errors_is",
            vec![arg_value(args, 0), arg_value(args, 1)],
        )),
        "errors.Join" => {
            if args.len() == 1 && args[0].spread {
                Some(go_builtin_call(
                    "__go_errors_join",
                    vec![args[0].value.clone()],
                ))
            } else {
                Some(go_builtin_call(
                    "__go_errors_join",
                    vec![go_array_of(
                        args.iter().map(|arg| arg.value.clone()).collect(),
                    )],
                ))
            }
        }
        _ => None,
    }
}

fn arg_value(args: &[Argument], idx: usize) -> Expression {
    args.get(idx)
        .map(|arg| arg.value.clone())
        .unwrap_or_else(Expression::null)
}

fn go_builtin_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
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
