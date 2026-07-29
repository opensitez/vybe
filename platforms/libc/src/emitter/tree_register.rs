//! `libc.*` namespace-tree registration for platform-owned libc surfaces.
//!
//! C also contributes profile-shaped libc entries, but platform math helpers
//! live here so any language can resolve `libc.math.*` without depending on
//! the C frontend being registered.

use std::sync::Once;

use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};

/// Register the platform libc surface under the `libc` root. Idempotent;
/// later C/profile registration merges with this tree.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut math = Subtree::new();
        for (name, emit) in [
            ("erf", "libc.math.erf"),
            ("erfc", "libc.math.erfc"),
            ("tgamma", "libc.math.tgamma"),
            ("gamma", "libc.math.tgamma"),
            ("lgamma", "libc.math.lgamma"),
        ] {
            math.insert(
                name.to_string(),
                NamespaceNode::CommonEmit(emit.to_string()),
            );
        }

        let mut root = Subtree::new();
        root.insert("math".to_string(), NamespaceNode::Namespace(math));
        namespaces::register_namespace_tree("libc", NamespaceNode::Namespace(root));
    });
}
