//! `dart.*` namespace-tree registration (namespaceplan.md Phase 5).
//!
//! Mirrors the dotnet registrar: the Dart language contributes DATA — its
//! profile `[builtins]` table and namespace aliases, the same data its
//! emit dispatch and linker execute — to the shared namespace tree.
//! Resolution LOGIC lives only in the common resolver; any language can
//! walk `dart.math.sqrt` (an alias leaf onto the canonical `ecma.math`).
//!
//! Leaf rules (dotnet template + the plan's alias-leaves):
//! - `common:dart.<fn>` emits register as `CommonEmit` leaves;
//! - host-backed builtins register as `Fn` leaves at their (dotted)
//!   builtin path (`dart.double.parse`);
//! - profile namespace aliases (`math` → `ecma:math`) register as
//!   `Alias` leaves onto the canonical tree path (`dart.math` →
//!   `ecma.math`) — source names reconciled to canonical names as data;
//! - opcode/intrinsic builtins have no process-global target — skipped.

use std::sync::Once;

use crate::emitter::namespaces::{self, NamespaceNode, Subtree};
use crate::profile::{BuiltinEmit, EsmDefault, parse_profile};

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

/// `"ecma:math"` → `"ecma.math"`, `"wasi:cli/terminal"` → `"wasi.cli.terminal"`.
fn module_tree_path(module: &str) -> String {
    module.replace([':', '/'], ".").to_lowercase()
}

/// Register the Dart surface under the `dart` root. Idempotent; first
/// call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let Ok(profile) = parse_profile(super::profile_source()) else {
            return;
        };
        let mut root = Subtree::new();
        for (name, def) in &profile.builtins {
            let key = name.to_lowercase();
            match &def.emit {
                BuiltinEmit::Common(op) => {
                    if let Some(path) = op.strip_prefix("dart.") {
                        insert_path(&mut root, path, NamespaceNode::CommonEmit(op.clone()));
                    }
                }
                BuiltinEmit::HostCall(module, func) => {
                    insert_path(&mut root, &key, namespaces::host_fn(module, func));
                }
                _ => {}
            }
        }
        for default in &profile.esm_defaults {
            if let EsmDefault::Namespace { alias, module } = default {
                insert_path(
                    &mut root,
                    &alias.to_lowercase(),
                    NamespaceNode::Alias(module_tree_path(module)),
                );
            }
        }
        namespaces::register_namespace_tree("dart", NamespaceNode::Namespace(root));
    });
}
