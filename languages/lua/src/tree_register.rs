//! `lua.*` namespace-tree registration.
//!
//! Lua contributes its builtin/library surface as resolver data. The compiler
//! resolves names through the shared namespace tree; Lua-specific behavior
//! remains in Lua-owned adapters reached through `CommonEmit` leaves.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};
use vybe_runtime::profile::{BuiltinEmit, parse_profile};

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let Some(leaf) = segments.pop() else {
        return;
    };
    if leaf.is_empty() {
        return;
    }

    let mut cursor = root;
    for segment in segments {
        if segment.is_empty() {
            return;
        }
        let entry = cursor
            .entry(segment.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.entry(leaf.to_string()).or_insert(node);
}

fn add_profile_leaf(root: &mut Subtree, name: &str, emit: &BuiltinEmit) {
    if name.starts_with("__") {
        return;
    }

    let node = match emit {
        BuiltinEmit::Common(op) => NamespaceNode::CommonEmit(op.clone()),
        BuiltinEmit::HostCall(module, func) => namespaces::host_fn(module, func),
        _ => return,
    };
    insert_path(root, name, node);
}

pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let Ok(profile) = parse_profile(crate::profile_source()) else {
            return;
        };

        let mut root: Subtree = BTreeMap::new();
        for (name, def) in &profile.builtins {
            add_profile_leaf(&mut root, name, &def.emit);
        }

        insert_path(
            &mut root,
            "dofile",
            NamespaceNode::CommonEmit("lua.dofile".to_string()),
        );

        namespaces::register_namespace_tree("lua", NamespaceNode::Namespace(root));
    });
}
