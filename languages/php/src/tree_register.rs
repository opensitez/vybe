//! `php.*` namespace-tree registration (namespaceplan.md: "PHP stdlib
//! surface (str_*, array_*, preg_*, …)", spelled like `common:*` emit
//! names).
//!
//! Mirrors the dotnet registrar: the PHP language contributes DATA — its
//! profile `[builtins]` table, the same table its emit dispatch executes —
//! to the shared namespace tree. Resolution LOGIC lives only in the
//! common resolver; any language can walk `php.str_replace`.
//!
//! Leaf rules (dotnet template):
//! - `common:php.<fn>` emits register as `CommonEmit` leaves at
//!   `php.<fn>`;
//! - host-backed builtins register as `Fn` leaves at `php.<name>`;
//! - opcode/intrinsic/print builtins have no process-global target to
//!   point at — skipped.

use std::sync::Once;

use vybe_emitter::namespaces::{self, NamespaceNode, Subtree};
use vybe_bytecode::profile::{BuiltinEmit, parse_profile};

/// Register the PHP stdlib surface under the `php` root. Idempotent;
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
            // Dotted builtin keys are member-shaped (receiver dispatch),
            // not free-function surface — the php.* package holds the
            // flat stdlib names only.
            if key.contains('.') {
                continue;
            }
            match &def.emit {
                BuiltinEmit::Common(op) => {
                    if op.strip_prefix("php.").is_some() {
                        root.entry(key)
                            .or_insert(NamespaceNode::CommonEmit(op.clone()));
                    }
                }
                BuiltinEmit::HostCall(module, func) => {
                    root.entry(key)
                        .or_insert_with(|| namespaces::host_fn(module, func));
                }
                _ => {}
            }
        }
        namespaces::register_namespace_tree("php", NamespaceNode::Namespace(root));
    });
}
