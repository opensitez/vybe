//! ⛔ KEYS KEEP THE DECLARED SPELLING. They used to be lowercased here, which
//! only worked because every tree lookup lowercased its query too. Lookups now
//! match EXACT first and fold only on a miss, so a case-sensitive language
//! resolves by the real name and a case-insensitive one still resolves by the
//! fold. See `documentation/casesensitivityplan.md`.
//! `flutter.*` namespace-tree registration.
//!
//! The Flutter platform contributes DATA — its widget catalog — to the shared
//! namespace tree in `vybe_runtime::namespaces`. Each widget registers as a
//! `Type` node at `flutter.<name>` carrying a `CtorSpec`, so the name is
//! reachable fully-qualified from ANY language and constructs through the ONE
//! common-resolver `Ctor` path (`Compiler::emit_tree_ctor_construction`) — no
//! Flutter-specific host functions, no per-module ctor globals.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_runtime::namespaces::{self, CtorSpec, FieldGui, NamespaceNode, Subtree};

use super::catalog::{self, FlutterClass};

/// Build the generic-construction spec for a widget: constructor params (in
/// declared order — positional slots already come first in the catalog),
/// the instance fields they capture (identical to the params), and the
/// `is`/`instanceof` ancestry (`self`-first parent chain plus any declared
/// interfaces/mixins).
fn ctor_spec(class: &FlutterClass) -> CtorSpec {
    let names: Vec<String> = class.fields.iter().map(|f| f.name.to_string()).collect();
    let mut ancestry: Vec<String> = catalog::ancestry(class);
    // **A widget IS an element** — so the identity chain says so, and stops
    // being two mechanisms.
    //
    // `Widget` is the catalog's root (`abstracts.rs`), and continuing it into
    // the live DOM type is what earns a real rtt: the shared ctor path allocates
    // with this chain, `reserve_platform_type` links each pair by
    // `parent_index`, and `ref.test` then answers `is StatelessWidget` /
    // `is Widget` / `is HTMLElement` natively instead of a `__types` string
    // scan. Method dispatch walks the same chain, so the DOM vtable stays
    // reachable from a widget.
    //
    // Only for a class with a backing control: `EdgeInsets` and `Color` are
    // data, they are not in the document, and nothing may make them elements.
    // The tail is also what the shared path GATES on, so declaring it here is
    // this platform opting in — dotnet and plib are untouched until they do the
    // same.
    if class.widget_host_fn.is_some() {
        ancestry.push(namespaces::DOM_ELEMENT_TYPE.to_string());
    }
    // ⚠ Interfaces come AFTER the DOM tail, and must: the chain is linked by
    // consecutive pairs, so a name inserted mid-ancestry would re-parent the
    // link it lands between. Every class carrying interfaces is a `data` type
    // (`data_with_interfaces`), which has no control and so no tail.
    for iface in class.interfaces {
        ancestry.push((*iface).to_string());
    }
    // Leaf controls whose child/label Text IS the control's own caption (a
    // button's face), rather than a nested child control.
    //
    // ⛔ RETIRED — and the name table with it. `FieldGui::Caption` set a control
    // PROPERTY from a widget, which stringifies: `AppBar(title: Text('x'))` and
    // `ElevatedButton(child: Text('7'))` both rendered the literal `[object]`
    // where the label belonged.
    //
    // It was true of the old control host, where a Button owned a `Text`
    // property and had
    // no children. HTML does not work that way and never did: a button's face
    // IS a child node — `<button><span>7</span></button>` — and so is an app
    // bar's title. Nesting is therefore not a special case of captioning, it is
    // the only case, and `NestOrProp` already decides it per VALUE (`ref.test`)
    // instead of per class NAME. A `Tab(text: 'x')` string still lands as a
    // property through the same arm, because the test asks what the value is.
    //
    // The list this replaced was one of the two name tables flexclassplan §4a
    // rules out: a widget added to the catalog tomorrow got the wrong role
    // silently unless someone remembered to extend a `matches!` here.
    let field_gui = class
        .fields
        .iter()
        .map(|f| {
            if f.children {
                FieldGui::Children
            } else if is_callback_field(f.name) {
                // **The DOM event the control actually fires**, not `click` for
                // everything. Choosing an option in a `<select>` fires `change`
                // — both engines agree, and neither has ever fired `click` for
                // it — so a `DropdownButton.onChanged` wired to `click` simply
                // never ran: picking an item changed the selection in the DOM
                // and the program was never told.
                //
                // The claim this replaced was that "the host routes control
                // events … to the Click handler". It does not: `drain_events`
                // maps `SelectChanged`/`DropdownSelected`/`ListBoxSelected` to
                // `change`, and webcore's form-event bridge does the same.
                //
                // A field may name its own event through `role` — a text field
                // wants `input`, which fires per keystroke as Flutter's
                // `onChanged` does, rather than `change` on commit.
                FieldGui::Event(
                    f.role.unwrap_or_else(|| dom_event_for(f.name)).to_string(),
                )
            } else if f.name == "data" {
                // `Text('x').data` is the control's caption.
                FieldGui::NestOrProp("Text".to_string())
            } else {
                // Semantic value fields (`value`/`min`/`max`/`text`/…) forward
                // as `Set<Prop>` commands to the backing control — under the
                // role the catalog DECLARES, which is the field's own name
                // unless it says otherwise (`MaterialApp.title` is the window
                // title, not a `title=""` tooltip).
                FieldGui::NestOrProp(f.role.unwrap_or(f.name).to_string())
            }
        })
        .collect();
    CtorSpec {
        params: names.clone(),
        fields: names,
        ancestry,
        control_fn: class.widget_host_fn.map(str::to_string),
        // No default chrome: a Flutter widget's children arrive as constructor
        // ARGUMENTS, so there is nothing a widget is born with that the call
        // site has not already said.
        inner_html: None,
        // Nor any construction-time part: a Flutter widget's whole content
        // arrives as arguments and is applied by the nesting path.
        after_create: None,
        // A Flutter widget argument may be a COMPOSITE — `home: CalculatorPage()`
        // is a `StatefulWidget`, a description with no element until `build()`
        // has run. `_vfConcrete` is that build, and it is the identity for
        // anything already concrete, so declaring it here costs a call and
        // decides nothing at compile time.
        nest_coerce: class
            .widget_host_fn
            .is_some()
            .then(|| "_vfConcrete".to_string()),
        field_gui,
        value_equality: is_value_equality_type(class.name),
    }
}

/// A user-supplied callback field (wires to the control's event, not a prop).
/// Flutter names all handlers `on…`: `onPressed`/`onChanged`/`onTap`/
/// `onLongPress`/`onDoubleTap`/`onSelected`/`onDeleted`/`onDismissed`/
/// `onPageChanged`/`onClosing`/`onSaved`.
/// The DOM event a Flutter callback field corresponds to.
///
/// Flutter names the INTENT (`onChanged`) and HTML names the EVENT (`change`),
/// and the two only coincide for a press. Anything not listed is a press:
/// `onPressed`, `onTap`, `onLongPress` are all `click`.
fn dom_event_for(field: &str) -> &'static str {
    match field {
        // A value the user picked or committed: `<select>`, `<input>` and
        // `<textarea>` all fire `change` for it.
        "onChanged" | "onSubmitted" | "onFieldSubmitted" | "onEditingComplete" => "change",
        _ => "click",
    }
}

fn is_callback_field(name: &str) -> bool {
    name.starts_with("on") && name.len() > 2 && name.as_bytes()[2].is_ascii_uppercase()
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
            | "EdgeInsets"
            | "EdgeInsetsDirectional"
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
                class.name.to_string(),
                NamespaceNode::Type {
                    ctor: Some(ctor_spec(class)),
                    ctor_call: None,
                    statics: Subtree::new(),
                    methods: BTreeMap::new(),
                    member_returns: std::collections::BTreeMap::new(),
                },
            );
        }
        namespaces::register_namespace_tree("flutter", NamespaceNode::Namespace(classes));
    });
}
