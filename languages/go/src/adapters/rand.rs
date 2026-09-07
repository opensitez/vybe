use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    insert_path(
        root,
        "rand.Intn",
        namespaces::host_fn("ecma:math", "random"),
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
