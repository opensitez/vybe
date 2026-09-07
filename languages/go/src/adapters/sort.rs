use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("sort.Ints", "collections.sort"),
        ("sort.Float64s", "collections.sort"),
        ("sort.Strings", "collections.sort"),
        ("sort.Search", "go.sort_search"),
        ("sort.Find", "go.sort_find"),
        ("sort.SearchInts", "go.sort_search_ordered"),
        ("sort.SearchStrings", "go.sort_search_ordered"),
        ("sort.SearchFloat64s", "go.sort_search_ordered"),
        ("sort.Slice", "go.sort_slice"),
        ("sort.SliceStable", "go.sort_slice"),
        ("sort.SliceIsSorted", "go.sort_is_sorted"),
        ("sort.Reverse", "go.sort_reverse"),
        ("sort.IntsAreSorted", "go.sort_is_sorted"),
        ("sort.Float64sAreSorted", "go.sort_is_sorted"),
        ("sort.StringsAreSorted", "go.sort_is_sorted"),
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
