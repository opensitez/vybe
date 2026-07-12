//! `plib.*` namespace-tree registration (namespaceplan.md: "Pascal GCL
//! surface").
//!
//! Mirrors the dotnet registrar: the plib platform contributes DATA — its
//! GCL class table (`gcl::gcl_classes()`, the same table the Pascal
//! lowering executes) — to the shared namespace tree in
//! `vybe_emitter::namespaces`. Resolution LOGIC lives only in the common
//! resolver; any language can walk `plib.tbutton.create`.
//!
//! Leaves follow the dotnet rules:
//! - each GCL class is a `Type` node at `plib.<class>`;
//! - its widget host constructor (`vybe:gui new_Button`, `newForm`, …)
//!   is the `create` static leaf — GCL's Delphi surface is literally
//!   `TButton.Create`;
//! - instance methods (Show/Close/Add) are receiver-dispatched, never
//!   tree-resolved — skipped, same as dotnet;
//! - chunk-built property accessors are per-compilation artifacts, not
//!   process-global surface — skipped, same as dotnet's `UserChunk`.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_emitter::gui;
use vybe_emitter::namespaces::{self, NamespaceNode, Subtree};

/// Register the GCL class table under the `plib` root. Idempotent; first
/// call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut classes = Subtree::new();
        for class in super::gcl::gcl_classes() {
            let mut statics = Subtree::new();
            if let Some(widget_fn) = class.widget_host_fn {
                statics.insert(
                    "create".to_string(),
                    namespaces::host_fn(gui::GUI_MODULE, widget_fn),
                );
            }
            classes.insert(
                class.name.to_lowercase(),
                NamespaceNode::Type {
                    ctor: None,
                    statics,
                    methods: BTreeMap::new(),
                },
            );
        }
        namespaces::register_namespace_tree("plib", NamespaceNode::Namespace(classes));
    });
}
