use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("bits.OnesCount", "go.bits_ones_count"),
        ("bits.OnesCount8", "go.bits_ones_count"),
        ("bits.OnesCount16", "go.bits_ones_count"),
        ("bits.OnesCount32", "go.bits_ones_count"),
        ("bits.OnesCount64", "go.bits_ones_count"),
        ("bits.LeadingZeros", "go.bits_leading_zeros"),
        ("bits.LeadingZeros8", "go.bits_leading_zeros"),
        ("bits.LeadingZeros16", "go.bits_leading_zeros"),
        ("bits.LeadingZeros32", "go.bits_leading_zeros"),
        ("bits.LeadingZeros64", "go.bits_leading_zeros"),
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
