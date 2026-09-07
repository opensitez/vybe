use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("path.Join", "go.path_join"),
        ("path.Clean", "go.path_clean"),
        ("path.Base", "go.path_base"),
        ("path.Ext", "go.path_ext"),
        ("path.IsAbs", "go.path_is_abs"),
        ("path.Split", "go.path_split"),
        ("filepath.Join", "go.path_join"),
        ("filepath.Clean", "go.path_clean"),
        ("filepath.Base", "go.path_base"),
        ("filepath.Ext", "go.path_ext"),
        ("filepath.IsAbs", "go.path_is_abs"),
        ("filepath.Split", "go.path_split"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
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
