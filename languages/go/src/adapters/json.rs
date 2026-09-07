use vybe_ast::{Argument, ExprKind, Expression};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("encoding.json.RawMessage", "__go_json_RawMessage"),
        ("encoding.json.Marshal", "__go_json_stringify"),
        ("encoding.json.MarshalIndent", "__go_json_stringify"),
        ("encoding.json.Unmarshal", "__go_json_parse"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }
    register_type(root, "encoding.json.RawMessage", &[]);
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "json.RawMessage" | "encoding/json.RawMessage" | "encoding.json.RawMessage" => {
            Some("__goRawMessage")
        }
        _ => None,
    }
}

pub(crate) fn rewrite_raw_message(args: &[Argument]) -> Option<Expression> {
    Some(Expression::new(ExprKind::Cast {
        expr: Box::new(
            args.first()
                .map(|arg| arg.value.clone())
                .unwrap_or_else(Expression::null),
        ),
        type_name: "__goRawMessage".to_string(),
    }))
}

fn register_type(root: &mut Subtree, name: &str, methods: &[&str]) {
    let mut method_tree = Subtree::new();
    for method in methods {
        method_tree.insert(
            (*method).to_string(),
            NamespaceNode::CommonEmit(format!("go.{name}.{method}")),
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
            member_returns: Default::default(),
        },
    );
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
