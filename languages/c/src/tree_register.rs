//! `libc.*` namespace-tree registration (namespaceplan.md: "C runtime
//! surface; header-driven mounts → `libc.*`").
//!
//! Mirrors the dotnet registrar: the C language contributes DATA — its
//! profile `[builtins]` table, the same table its emit dispatch executes —
//! to the shared namespace tree. Resolution LOGIC lives only in the common
//! resolver; any language can walk `libc.stdio.printf`.
//!
//! Leaf rules (dotnet template):
//! - `common:libc.<hdr>.<fn>` emits register at their own emit path as
//!   `CommonEmit` leaves (`libc.stdio.printf`);
//! - host-backed builtins register as `Fn` leaves at `libc.<name>`;
//! - opcode/intrinsic/print builtins have no process-global target to
//!   point at — skipped, same as dotnet's chunk-backed methods.

use std::sync::Once;

use vybe_emitter::namespaces::{self, NamespaceNode, Subtree};
use vybe_plugin::profile::{BuiltinEmit, parse_profile};

/// Insert `node` at the dotted `path` under `root`, creating interior
/// namespaces as needed. Keys are lowercase-canonical.
fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let leaf = segments.pop().expect("non-empty path");
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return; // leaf/namespace collision: first registration wins
        };
        cursor = children;
    }
    cursor.entry(leaf.to_string()).or_insert(node);
}

/// Register the C runtime surface under the `libc` root. Idempotent;
/// first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let Ok(profile) = parse_profile(super::profile_source()) else {
            return;
        };
        let mut root = Subtree::new();
        for (name, def) in &profile.builtins {
            match &def.emit {
                BuiltinEmit::Common(op) => {
                    if let Some(path) = op.strip_prefix("libc.") {
                        insert_path(&mut root, path, NamespaceNode::CommonEmit(op.clone()));
                    }
                }
                BuiltinEmit::HostCall(module, func) => {
                    insert_path(
                        &mut root,
                        &name.to_lowercase(),
                        namespaces::host_fn(module, func),
                    );
                }
                _ => {}
            }
        }
        namespaces::register_namespace_tree("libc", NamespaceNode::Namespace(root));
    });
}
