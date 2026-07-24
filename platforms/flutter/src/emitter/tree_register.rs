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
    // Leaf controls whose child/label Text IS the control's own caption (a
    // button's face), rather than a nested child control.
    let is_captioned_leaf = matches!(
        class.name,
        "ElevatedButton"
            | "TextButton"
            | "OutlinedButton"
            | "CupertinoButton"
            | "FloatingActionButton"
            | "Chip"
            | "ActionChip"
            | "FilterChip"
            | "ChoiceChip"
    );
    let field_gui = class
        .fields
        .iter()
        .map(|f| {
            let caption = (is_captioned_leaf && (f.name == "child" || f.name == "label"))
                || (class.name == "AppBar" && f.name == "title")
                || (class.name == "Tab" && f.name == "text");
            if caption {
                // A leaf control (button/app-bar) takes its child Text as its
                // own caption, not as a nested child.
                FieldGui::Caption
            } else if f.children {
                FieldGui::Children
            } else if is_callback_field(f.name) {
                // Every user callback (`onPressed`/`onChanged`/`onTap`/…) wires
                // to the control's event; the host routes control events
                // (ButtonClicked/CheckboxToggled/TextChanged/SelectChanged) to
                // the "Click" handler.
                FieldGui::Event("Click".to_string())
            } else if f.name == "data" {
                // `Text('x').data` is the control's caption.
                FieldGui::NestOrProp("Text".to_string())
            } else {
                // Semantic value fields (`value`/`min`/`max`/`text`/…) forward
                // as `Set<Prop>` commands to the backing control.
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
        value_equality: is_value_equality_type(class.name),
    }
}

/// A user-supplied callback field (wires to the control's event, not a prop).
/// Flutter names all handlers `on…`: `onPressed`/`onChanged`/`onTap`/
/// `onLongPress`/`onDoubleTap`/`onSelected`/`onDeleted`/`onDismissed`/
/// `onPageChanged`/`onClosing`/`onSaved`.
fn is_callback_field(name: &str) -> bool {
    name.starts_with("on")
        && name.len() > 2
        && name.as_bytes()[2].is_ascii_uppercase()
}

/// Flutter immutable value types whose `operator ==` is by VALUE (structural),
/// not identity. Keys are the classic case (`ValueKey('a') == ValueKey('a')`);
/// so are the small geometry/paint value classes. Identity types (`GlobalKey`,
/// `UniqueKey`, `FocusNode`, controllers, notifiers) are deliberately EXCLUDED —
/// two distinct instances are never equal.
fn is_value_equality_type(name: &str) -> bool {
    matches!(
        name,
        "Key"
            | "ValueKey"
            | "ObjectKey"
            | "GlobalObjectKey"
            | "Color"
            | "Offset"
            | "FractionalOffset"
            | "Size"
            | "Radius"
            | "IconData"
            | "TextStyle"
    )
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
