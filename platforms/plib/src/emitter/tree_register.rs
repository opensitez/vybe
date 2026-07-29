//! `plib.*` namespace-tree registration (namespaceplan.md: "Pascal GCL
//! surface").
//!
//! Mirrors the dotnet registrar: the plib platform contributes DATA — its
//! GCL class table (`gcl::gcl_classes()`, the same table the Pascal
//! lowering executes) — to the shared namespace tree in
//! `vybe_bytecode::namespaces`. Resolution LOGIC lives only in the common
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

use vybe_compiler::primitives::gui;
use vybe_bytecode::namespaces::{self, CtorSpec, FieldGui, NamespaceNode, Subtree};

/// The `is`/`inherits` ancestry of a GCL class — self first, then its parent
/// chain, so `TButton is TControl` answers from the same `__types` stamp every
/// other adapter uses.
fn ancestry(class: &super::gcl::GclClass) -> Vec<String> {
    let mut out = vec![class.name.to_string()];
    let mut parent = class.parent;
    while let Some(name) = parent {
        out.push(name.to_string());
        parent = super::gcl::gcl_classes()
            .iter()
            .find(|c| c.name.eq_ignore_ascii_case(name))
            .and_then(|c| c.parent);
    }
    out
}

/// The generic-construction spec for a GCL class — the SAME shape Flutter
/// registers. `control_fn` is the `vybe:gui` control factory; each declared
/// property forwards as a `Set<Prop>` command on that control. This is what
/// makes plib an adapter over the one GUI interface instead of a compiler-side
/// class-registration pass.
fn ctor_spec(class: &super::gcl::GclClass) -> CtorSpec {
    let fields: Vec<String> = class.properties.iter().map(|p| p.to_string()).collect();
    CtorSpec {
        params: fields.clone(),
        field_gui: fields
            .iter()
            .map(|f| FieldGui::NestOrProp(f.clone()))
            .collect(),
        fields,
        ancestry: ancestry(class),
        control_fn: class.widget_host_fn.map(str::to_string),
        value_equality: false,
    }
}

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
                    ctor: Some(ctor_spec(class)),
                    ctor_call: None,
                    statics,
                    methods: BTreeMap::new(),
                member_returns: std::collections::BTreeMap::new(),
                },
            );
        }
        namespaces::register_namespace_tree("plib", NamespaceNode::Namespace(classes));
    });
}
