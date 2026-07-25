//! `java.*` namespace-tree registration (namespaceplan.md Phase 5 shape).
//!
//! Mirrors the dotnet registrar: the Java language contributes DATA — its
//! profile `[builtins]` table, the same table its emit dispatch executes —
//! to the shared namespace tree. Resolution LOGIC lives only in the
//! common resolver; any language can walk `java.integer.parseint`.
//!
//! Leaf rules (dotnet template):
//! - Java package-surface common emits register as `CommonEmit` leaves at
//!   the builtin's own (dotted) key path (`java.util.Objects.equals` →
//!   `java.util.objects.equals`), even when the actual common op is a
//!   shared category such as `object.equals`;
//! - Java shorthand common emits (`Integer.parseInt`) register when the
//!   target is Java-owned (`common:java.<op>`);
//! - host-backed builtins register as `Fn` leaves at their key path;
//! - opcode/intrinsic/print builtins have no process-global target to
//!   point at — skipped.

use std::sync::Once;

use vybe_emitter::namespaces::{self, NamespaceNode, Subtree};
use vybe_bytecode::profile::{BuiltinEmit, parse_profile};

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

/// Register the Java stdlib surface under the `java` root. Idempotent;
/// first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let Ok(profile) = parse_profile(super::profile_source()) else {
            return;
        };
        let mut root = Subtree::new();
        for (name, def) in &profile.builtins {
            let key = name.to_lowercase();
            // Internal walker-support helpers are not surface.
            if key.starts_with("__") {
                continue;
            }
            match &def.emit {
                BuiltinEmit::Common(op) => {
                    if key.starts_with("java.") || op.starts_with("java.") {
                        insert_path(&mut root, &key, NamespaceNode::CommonEmit(op.clone()));
                    }
                }
                BuiltinEmit::HostCall(module, func) => {
                    insert_path(&mut root, &key, namespaces::host_fn(module, func));
                }
                _ => {}
            }
        }
        namespaces::register_namespace_tree("java", NamespaceNode::Namespace(root));
    });
}
