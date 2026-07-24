//! `flutter.*` namespace-tree registration.
//!
//! The Flutter platform contributes DATA — its widget catalog — to the shared
//! namespace tree in `vybe_emitter::namespaces`. Each widget registers as a
//! `Type` node at `flutter.<name>` carrying a `CtorSpec`, so the name is
//! reachable fully-qualified from ANY language and constructs through the ONE
//! common-resolver `Ctor` path (`Compiler::emit_tree_ctor_construction`) — no
//! Flutter-specific host functions, no per-module ctor globals.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_emitter::namespaces::{self, CtorSpec, FieldGui, NamespaceNode, Subtree};

use super::catalog::{self, FlutterClass};

/// Build the generic-construction spec for a widget: constructor params (in
/// declared order — positional slots already come first in the catalog),
/// the instance fields they capture (identical to the params), and the
/// `is`/`instanceof` ancestry (`self`-first parent chain plus any declared
/// interfaces/mixins).
fn ctor_spec(class: &FlutterClass) -> CtorSpec {
    let names: Vec<String> = class.fields.iter().map(|f| f.name.to_string()).collect();
    let mut ancestry: Vec<String> = catalog::ancestry(class)
        .into_iter()
        .map(str::to_string)
        .collect();
    for iface in class.interfaces {
        ancestry.push((*iface).to_string());
    }
    let field_gui = class
        .fields
        .iter()
        .map(|f| {
            let caption = (class.name == "ElevatedButton" && f.name == "child")
                || (class.name == "AppBar" && f.name == "title");
            if caption {
                // A leaf control (button/app-bar) takes its child Text as its
                // own caption, not as a nested child.
                FieldGui::Caption
            } else if f.children {
                FieldGui::Children
            } else if f.name == "onPressed" || f.name == "onLongPress" {
                FieldGui::Event("Click".to_string())
            } else if f.name == "data" {
                // `Text('x').data` is the control's caption.
                FieldGui::NestOrProp("Text".to_string())
            } else {
                FieldGui::NestOrProp(f.name.to_string())
            }
        })
        .collect();
    CtorSpec {
        params: names.clone(),
        fields: names,
        ancestry,
        control_fn: class.widget_host_fn.map(str::to_string),
        field_gui,
    }
}

/// Register the Flutter widget catalog under the `flutter` root. Idempotent;
/// first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut classes = Subtree::new();
        for class in catalog::flutter_classes() {
            classes.insert(
                class.name.to_lowercase(),
                NamespaceNode::Type {
                    ctor: Some(ctor_spec(class)),
                    statics: Subtree::new(),
                    methods: BTreeMap::new(),
                },
            );
        }
        namespaces::register_namespace_tree("flutter", NamespaceNode::Namespace(classes));
    });
}
