use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("os.ReadFile", "filesystem.read_file"),
        ("os.WriteFile", "filesystem.write_file"),
        ("os.Mkdir", "filesystem.mkdir"),
        ("os.Remove", "filesystem.remove"),
        ("os.Exit", "control_flow.exit"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    insert_path(root, "os.Getenv", namespaces::host_fn("wasi:cli", "getEnv"));
    insert_path(root, "os.Setenv", namespaces::host_fn("wasi:cli", "setEnv"));
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
