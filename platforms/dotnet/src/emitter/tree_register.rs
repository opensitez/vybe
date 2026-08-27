//! ⛔ KEYS KEEP THE .NET DECLARED SPELLING (`Console`, `WriteLine`, `Length`).
//!
//! They used to be lowercased here. That served VB — which is case-insensitive
//! — and quietly mis-served C#, which is not: a rule that folds is wrong on
//! C#'s own terms even when it happens to answer. Tree lookups now match EXACT
//! first and fold only on a miss, so C# resolves by the real spelling and VB
//! still resolves by the fold. One registration, both languages, neither
//! compromised. See `documentation/casesensitivityplan.md`.
//! `dotnet.*` namespace-tree registration (namespaceplan.md, dotnet phase).
//!
//! The dotnet platform contributes DATA — its component-model class
//! descriptors — to the shared namespace tree in `vybe_runtime::namespaces`.
//! Resolution LOGIC lives only in the common resolver: VB, C#, and every
//! other language resolve `dotnet.system.console.writeline` through the
//! same tree walk, instead of a platform-owned dotted-name cascade
//! (`resolver.rs`, which this registration supersedes and which dissolves
//! once VB/C# routing is fully migrated).

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_runtime::component_model::{ConstructorTarget, MethodBody};
use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};

/// Register every component class descriptor as a `Type` node at
/// `<interface path>.<class name>` — statics become `CommonEmit`/host-fn
/// leaves (the same `MethodBody` targets `dispatch.rs` executes).
/// Idempotent; first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        for export in super::class_exports::dotnet_class_exports() {
            let (interface, class) = (export.interface, &export.class);
            let mut statics = Subtree::new();
            // INSTANCE members, registered as real target nodes. They used to be
            // skipped because `Type.methods` was a bare `FuncSig` map that could
            // not carry a target — so the compiler reached into this crate
            // directly to look them up. `methods` is a `Subtree` now, so a
            // platform declares its instance surface here like everything else.
            // Overloads collected per name in DECLARATION order — a .NET
            // class routinely declares `Reverse()` and `Reverse(i, n)` to
            // different targets, and a name-keyed map keeps only one of them.
            let mut method_overloads: BTreeMap<String, Vec<(u8, NamespaceNode)>> = BTreeMap::new();
            let mut static_overloads: BTreeMap<String, Vec<(u8, NamespaceNode)>> = BTreeMap::new();
            let mut methods = Subtree::new();
            // THE control fact, decided ONCE per class and consumed by every
            // claim below (method verbs, `Controls` seeding, the ctor spec).
            // Being a control is a WINDOWS.FORMS fact, never a leaf-name
            // fact: `System.Threading.Timer` declares a `Dispose` too, and a
            // leaf-keyed answer registered it as `gui.ctrl.dispose` — whose
            // emitter touches `activeDocument`, which CREATES a document,
            // which made a headless callback-timer program present a real
            // window and park in the GUI event loop. Three copies of this
            // gate briefly existed; copies of one rule drift, so the fact is
            // computed here and only here.
            let class_is_control = interface
                .eq_ignore_ascii_case("dotnet.System.Windows.Forms")
                && element_backed_control(&class.name);
            // A form is `document.body`, so node-scoped destruction does not
            // apply to it — see `gui_control_verb`. `inherits` walks the
            // descriptor chain, so `class MyForm : Form` answers true too.
            let class_is_form = namespaces::inherits(&class.name, "Form", descriptor_parent);
            // The WHOLE inherited surface, not just this class's own methods —
            // the tree has no parent link, so anything left off here is
            // unreachable. See `inherited_methods`.
            for m in &inherited_methods(&class.name) {
                // A control METHOD is a shared VERB, resolved through
                // `primitives/gui.rs` like its properties — never a host call.
                // Gated on the class actually being a control for
                // the same reason the accessors are: `Graphics` and the value
                // types declare methods too, and `Update`/`Refresh` are
                // ordinary names that must not be captured off a non-control.
                let gui_verb = if class_is_control && !m.is_static {
                    gui_control_verb(&m.name, class_is_form)
                } else {
                    None
                };
                if let Some(verb) = gui_verb {
                    // Buckets into `method_overloads` exactly like every other
                    // method rather than inserting straight into `methods`,
                    // because the ARITY has to travel with the node: a method
                    // call resolves by (name, arity), so a bare `CommonEmit`
                    // leaf is found as a name and then not called — `b.Hide()`
                    // reached `undefined is not callable`. Same registration
                    // path, only the target differs.
                    let entries = method_overloads.entry(m.name.to_string()).or_default();
                    // Same "first declaration of an arity wins" rule as the
                    // ordinary path below. `inherited_methods` yields nearest
                    // first, so an override shadows the base declaration it
                    // re-declares instead of appending a second entry that can
                    // never be selected.
                    if !entries.iter().any(|(a, _)| *a == m.arity) {
                        entries.push((
                            m.arity,
                            NamespaceNode::CommonEmit(format!(
                                "{}{verb}",
                                vybe_compiler::primitives::gui::CTRL_METHOD_EMIT
                            )),
                        ));
                    }
                    continue;
                }
                let node = match &m.body {
                    MethodBody::Common(emit) => NamespaceNode::CommonEmit(emit.clone()),
                    // The descriptor knows the arity — record it, so the
                    // compiler can select by arity from the tree instead of
                    // calling into this crate.
                    MethodBody::HostCall(t) => {
                        namespaces::host_fn_with_arity(&t.module, &t.name, m.arity)
                    }
                    // Chunk-backed methods are per-compilation artifacts,
                    // not process-global surface.
                    MethodBody::UserChunk(_) => continue,
                };
                let bucket = if m.is_static {
                    &mut static_overloads
                } else {
                    &mut method_overloads
                };
                let entries = bucket.entry(m.name.to_string()).or_default();
                // First declaration of an arity wins, matching the
                // descriptor-order scan this registration replaces.
                if !entries.iter().any(|(a, _)| *a == m.arity) {
                    entries.push((m.arity, node));
                }
            }
            for (name, entries) in method_overloads {
                methods.insert(name, namespaces::overloads(entries));
            }
            for (name, entries) in static_overloads {
                statics.insert(name, namespaces::overloads(entries));
            }
            // Properties, as real property members carrying BOTH directions.
            // Walked up the parent chain and flattened, because the tree has no
            // parent link and `Button.Text` is declared on `Control` — a flat
            // registration would leave every inherited property unreachable.
            // Nearest declaration wins, so an override shadows its base.
            let mut property_returns: BTreeMap<String, String> = BTreeMap::new();
            let is_control = element_backed_control(&class.name);
            // ANYTHING WITHOUT AN ELEMENT DECLARES NO PROPERTY ACCESSOR, so its
            // properties resolve the way a user-declared class's do: an ordinary
            // STRUCT FIELD read.
            //
            // That covers the `System.Drawing` value types (`Point.X`) and the
            // NON-VISUAL components (`Timer.Interval`, `BindingSource.DataSource`)
            // alike, and it is not a behaviour change: both bound the two
            // generic property host functions, and for an object with no
            // element that host call collapsed to exactly a lowercased field
            // access — the getter returned `o.properties.get(&prop_lower)` and
            // the setter did `properties.insert(prop_lower, val)`. Those ARE the field names
            // `emit_value_type_new` stores, so declaring nothing reproduces the
            // behaviour with no host at all.
            //
            // `primitives/gui.rs::declared_property_role` records the other half:
            // a class with no platform role gets a struct field read, which is
            // why a user-declared `Class Point` works and why the platform
            // `Point.X` role capturing it looked like object aliasing.
            for p in inherited_properties(&class.name) {
                if !is_control {
                    continue;
                }
                let node = namespaces::property(
                    p.getter
                        .as_ref()
                        .map(|t| accessor_node(t, &p.name, is_control, &class.name)),
                    p.setter
                        .as_ref()
                        .map(|t| accessor_node(t, &p.name, is_control, &class.name)),
                );
                // What the property READS BACK as, declared from its ROLE so
                // the one answer serves every frontend registered this way.
                // Not decoration: a property whose type is undeclared reads as
                // `null` to the expression machinery, so `btn.Text.Length` or
                // `if (c.Enabled)` had nothing to work on.
                if is_control {
                    if let Some(value_type) = vybe_compiler::primitives::gui::property_value_type(
                        match gui_property_role(&p.name) {
                            "" => p.name.to_ascii_lowercase(),
                            r => r.to_string(),
                        }
                        .as_str(),
                    ) {
                        property_returns.insert(p.name.to_string(), value_type.to_string());
                    }
                }
                methods
                    .entry(p.name.to_string())
                    .or_insert_with(|| node.clone());
                statics.entry(p.name.to_string()).or_insert(node);
            }
            for (name, node) in shared_emit_accessors(&class.name) {
                methods.insert(name, node);
            }
            if interface.eq_ignore_ascii_case("dotnet.System")
                && class.name.eq_ignore_ascii_case("Console")
            {
                // `Console.Out` / `Console.Error` — the CLR spelling. These were
                // authored lowercase for a tree that folded everything; the
                // fold is gone, so the DECLARATION has to be right.
                statics.insert("Out".into(), console_stdout_writer_node());
                statics.insert("Error".into(), console_stderr_writer_node());
            }
            if interface.eq_ignore_ascii_case("dotnet.System")
                && class.name.eq_ignore_ascii_case("Object")
            {
                statics.insert(
                    "Equals".into(),
                    NamespaceNode::CommonEmit("dotnet.object_equals".into()),
                );
                statics.insert(
                    "ReferenceEquals".into(),
                    NamespaceNode::CommonEmit("dotnet.object_reference_equals".into()),
                );
            }
            if interface.eq_ignore_ascii_case("dotnet.System")
                && class.name.eq_ignore_ascii_case("DateTime")
            {
                statics.insert(
                    "MinValue".into(),
                    NamespaceNode::CommonEmit("dotnet.datetime_min_value".into()),
                );
                statics.insert(
                    "MaxValue".into(),
                    NamespaceNode::CommonEmit("dotnet.datetime_max_value".into()),
                );
            }
            // Declare return types with the class, so the compiler reads them
            // from the tree instead of calling a dotnet-side name cascade.
            let mut member_returns = std::collections::BTreeMap::new();
            for m in &class.methods {
                // Statics and instance members alike — the compiler asks the
                // tree for a member's declared return type and must not care
                // which kind it is. Neither dotnet-side resolver takes an
                // arity, so a name key is exactly equivalent.
                let rt = if m.is_static {
                    super::static_method_return_type(&class.name, &m.name).map(str::to_string)
                } else {
                    super::instance_method_return_type(&class.name, &m.name)
                };
                if let Some(rt) = rt {
                    member_returns.insert(m.name.to_string(), rt);
                }
            }
            // Property roles fill in what the method scan does not cover. A
            // method's declared return wins on a name collision — it was
            // resolved from the descriptor's own signature, which is the more
            // specific statement.
            for (name, ty) in property_returns {
                member_returns.entry(name).or_insert(ty);
            }
            // A PROPERTY's declared type, for the classes that are not
            // controls. The scan above only walks `class.methods`, so a
            // property read came back untyped and the next hop on it resolved
            // against nothing — `cmd.Parameters.AddWithValue(…)` died there
            // even with the collection declared, because `Parameters` itself
            // said nothing about what it is.
            for (name, ty) in super::declared_instance_property_types(&class.name) {
                member_returns
                    .entry(name.to_string())
                    .or_insert_with(|| (*ty).to_string());
            }
            // A member that IS the receiver overrides both — the descriptor's
            // own signature describes the collection wrapper WinForms models,
            // and the element does not have one.
            for (name, ty) in self_member_returns(&class.name) {
                member_returns.insert((*name).to_string(), (*ty).to_string());
            }
            // `Controls` reads back as a control, so the NEXT hop resolves
            // `Add` against `Control`'s node — where `shared_emit_accessors`
            // declares it. Without this the getter answers the element and the
            // member after it resolves against nothing, which is the exact
            // failure `self_member_returns` was written for on `Items`.
            if class_is_control {
                // `Control.Controls` — the declared property name.
                member_returns.insert("Controls".to_string(), "Control".to_string());
            }

            // The descriptor's backing constructor, as a tree node. dotnet
            // classes are not generic field-capture constructions (no
            // `CtorSpec`) — `new Dictionary()` is a host factory call, so it
            // registers here in the same vocabulary as every other member.
            let ctor_call = class
                .constructor
                .as_ref()
                .and_then(|c| c.backing.as_ref())
                .map(|backing| {
                    Box::new(match backing {
                        ConstructorTarget::Host(t) => namespaces::host_fn(&t.module, &t.name),
                        ConstructorTarget::Common(emit) => NamespaceNode::CommonEmit(emit.clone()),
                    })
                });

            // A converted control DECLARES the element it is, the same way
            // plib and flutter do — which is what finally lets
            // `registered_control_element` answer for dotnet instead of
            // falling back to a `<vybe-*>` custom element.
            //
            // Registered ONLY for classes with a real HTML counterpart, so an
            // unconverted control keeps today's path exactly. A spec with
            // `control_fn` set routes construction through
            // `emit_tree_ctor_construction`, which creates the element AND
            // stamps `__type`/`__types` — identity these controls never had.
            // That is safe now only because properties and methods resolve
            // from the TREE (see `accessor_node` above and the `MethodBody`
            // registration); while they lived on ctor-bound thunks, taking
            // this path would have dropped them.
            // Consumes `class_is_control` — the one carrier of the fact.
            let ctor = class_is_control
                .then(|| html_element_for_control(&class.name))
                .flatten()
                .map(|element| control_ctor_spec(&class.name, element));

            let ty = NamespaceNode::Type {
                ctor,
                ctor_call,
                statics,
                methods,
                member_returns,
            };

            // "dotnet.System" + "Math" → dotnet.system.math
            //
            // A name with N segments becomes N nodes and nothing else. A class
            // is NEVER also registered under its own bare name as a top-level
            // root: `Math`, `Console` and `Object` used to be registered
            // alongside `dotnet` itself, which put every `System` type in the
            // same flat space as the user's own declarations and let a user
            // class named `Point` or `Component` collide with a platform type.
            //
            // Bare `Console.WriteLine` still resolves — through the profile,
            // which is where that decision belongs: `csharp`, `vb` and
            // `powershell` each declare `[[esm_default]] kind = "tree-ambient"
            // path = "dotnet.system"`, so an unqualified name is looked up
            // under that namespace by the common resolver. A language that
            // wants the short spelling says so in its profile instead of
            // inheriting it from a registration side effect.
            let mut segments: Vec<String> =
                interface.split('.').map(|s| s.to_string()).collect();
            segments.push(class.name.to_string());

            let mut node = ty;
            while segments.len() > 1 {
                let key = segments.pop().expect("non-empty");
                let mut children = Subtree::new();
                children.insert(key, node);
                node = NamespaceNode::Namespace(children);
            }
            namespaces::register_namespace_tree(&segments.pop().expect("root"), node);
        }
        register_color_statics();
    });
}

/// `Color.Red`, `Color.White`, … — the named colour statics.
///
/// `Color` reaches the tree through `constructor_class(…, "Color", …)`, which
/// declares a CONSTRUCTOR and no properties, so its `statics` arrived empty and
/// `Red` was not a leaf at all. `Color.White` then resolved to nothing and the
/// chain fell back to a runtime read of the bare root name — `global.get
/// system`, a global that does not exist — so every colour came out `null` and
/// painted `#00000000`. Silent, because a missed leaf under a registered root
/// reads as null rather than failing.
///
/// Registered as PATH, one node per segment: `dotnet` → `system` → `drawing` →
/// `color` → `red`. `merge_into` creates the levels that are missing and reuses
/// the ones that exist, and because `color` is already a `Type`, its
/// `Type{statics} × Namespace` arm folds these in as that type's statics.
/// Registration order does not matter — the mirror arm handles the reverse.
fn register_color_statics() {
    let mut colors = Subtree::new();
    for (member, _) in super::classes::drawing::COLOR_STATICS {
        colors.insert(
            member.to_string(),
            NamespaceNode::CommonEmit(super::classes::drawing::color_static_emit_key(member)),
        );
    }

    let mut node = NamespaceNode::Namespace(colors);
    // ⛔ The DECLARED spelling, and it must match what the main registration
    // loop produces (`interface.split('.')` → `System`, `Drawing`, plus
    // `class.name` → `Color`). These merge into that type's `statics` only if
    // the KEYS ARE THE SAME; a lowercase path builds a SECOND branch instead,
    // and `Color.Red` then resolves only by folding.
    for segment in ["Color", "Drawing", "System"] {
        let mut parent = Subtree::new();
        parent.insert(segment.to_string(), node);
        node = NamespaceNode::Namespace(parent);
    }
    namespaces::register_namespace_tree("dotnet", node);
}

/// A WinForms property spelling → the shared GUI ROLE it fills.
///
/// **A property is a role, not a name.** This function is the whole of dotnet's
/// job here: the WinForms word stops at this line and nothing downstream knows
/// WinForms exists — the same contract `platforms/plib` states for VCL, where
/// `Caption` and `Text` are one role.
///
/// Most WinForms spellings ARE the canonical role already (`Text`, `Name`,
/// `Enabled`, `Visible`, `Left`/`Top`/`Width`/`Height`), which is why this is
/// close to the identity: the shared property vocabulary was modelled on
/// WinForms in the first place. Only the names that disagree are listed.
///
/// Everything unlisted keeps its own spelling and lands on an ATTRIBUTE, which
/// is where unknown properties belong on the web. That is the correct
/// destination, not a gap: `Anchor`, `Dock`, `BackColor`, `Font`, `Cursor` and
/// the rest have no IDL counterpart, and inventing a role for each would grow
/// the shared table per control — exactly what `property_op` is written to
/// avoid.
fn gui_property_role(prop: &str) -> &'static str {
    match prop.to_ascii_lowercase().as_str() {
        // `Text` is the text role for EVERY control: the widget resolves what
        // it means for the element it is — a caption on a Label, the value on
        // a TextBox — so no element test belongs here.
        "text" | "caption" => "text",
        "name" => "name",
        // CheckBox / RadioButton. `CheckState` is the tri-state spelling of
        // the same role; the widget reads its value.
        "checked" | "checkstate" => "checked",
        "enabled" => "enabled",
        "visible" => "visible",
        "left" => "left",
        "top" => "top",
        "width" | "clientwidth" => "width",
        "height" | "clientheight" => "height",
        // A `Point`/`Size` value IS the pair of declarations above, and the
        // shared write path decomposes it. Designer-generated forms use these
        // and never `Left`/`Top`, so without the role every `.Designer.vb`
        // control sat at the origin with `location="[object]"` in the markup.
        "location" => "location",
        "size" | "clientsize" => "size",
        // ComboBox / ListBox selection, the same role plib maps `ItemIndex` to.
        // Style properties. Every role below ALREADY existed in the shared
        // table and was reachable from plib — VCL's `Color` maps to
        // `backcolor`, `Align` to `dock` — but WinForms had no row, so each
        // one fell through to an ATTRIBUTE of the same name and reached
        // nothing. `.dfm`/`.Designer.cs` forms set these constantly.
        "font" => "font",
        "backcolor" => "backcolor",
        "forecolor" => "forecolor",
        "dock" => "dock",
        "padding" => "padding",
        "margin" => "margin",
        "cursor" => "cursor",
        "tabindex" => "tabindex",
        "backgroundimage" => "backgroundimage",
        "selectedindex" => "selectedindex",
        "items" => "items",
        "readonly" => "readonly",
        "maxlength" => "maxlength",
        _ => "",
    }
}

/// A WinForms control METHOD → the shared GUI verb it IS.
///
/// The mirror of `gui_property_role`, and the same contract: the WinForms word
/// stops here. `Show`/`Hide`/`Focus`/`BringToFront` are the same verbs the VCL
/// spells on `TControl`, so both frontends reach one implementation in
/// `primitives/gui.rs` instead of each calling its own host fn.
///
/// `Refresh`/`Invalidate`/`Update` map to real verbs that lower to NOTHING —
/// a document repaints itself, so there is nothing for an author to ask for.
/// That is deliberate and is explained at `emit_gui_control_method`.
///
/// `is_form` exists for ONE verb. `Dispose` detaches the node, and a form is
/// mapped to `body` — so disposing a form would remove the document body and
/// take the whole app shell with it. A form's own lifetime is a WINDOW
/// question, not a node question, and it stays on its existing route until
/// forms stop being the body.
fn gui_control_verb(method: &str, is_form: bool) -> Option<&'static str> {
    Some(match method.to_ascii_lowercase().as_str() {
        "show" => "show",
        "hide" => "hide",
        "focus" => "focus",
        // `Control.Select()` IS `Focus()` — `control.rs` already declared both
        // against the same `__ctrl_focus`, so this is the existing alias said
        // once more on the DOM route rather than a new behaviour.
        "select" => "focus",
        "refresh" => "refresh",
        "invalidate" => "invalidate",
        "update" => "update",
        // Z-order IS document order, so raising a control is re-appending it
        // to its parent: `parentNode` + `appendChild`, both registered on
        // `web:dom`. `emit_gui_control_verb` already implements it.
        "bringtofront" => "bring_to_front",
        // `SendToBack` is `insertBefore` against the parent's current first
        // child. It was held back only because `web:dom` had no `firstChild`;
        // that is registered now, so both halves of the z-order pair lower to
        // the DOM.
        "sendtoback" => "send_to_back",
        // `Dispose` DESTROYS the control: `ChildNode.remove()`. The host fn it
        // replaces only set `Visible=false` and dropped the handler table, so a
        // "disposed" control could be brought back by writing `Visible` — which
        // is `Hide`, a different verb with a different promise.
        "dispose" if !is_form => "dispose",
        _ => return None,
    })
}

/// One accessor leaf.
///
/// A control's accessors bind to the shared role emits
/// (`gui.prop_get.<role>` / `gui.prop_set.<role>`) that `primitives/gui.rs`
/// lowers to `web:dom` / `web:html` / `web:cssom` — the same target plib
/// reaches. Both drive `widgets` underneath; the route is a DOM operation.
///
/// A dedicated per-property host function (`Environment.NewLine` →
/// `node:os.EOL`) is NOT a control property and is left exactly as it was.
///
/// `is_control` is what keeps the value types out. `Point`/`Size`/`Font` bind
/// the same two generic host functions, so the target name alone cannot tell
/// them apart from a `Button` — only the class can, and it does it by whether
/// it descends from `Control`.
/// Does this control INHERIT `Text` without ever painting it?
///
/// `Control.Text` exists on everything; drawing it is the individual control's
/// business. A `Label` or `Button` is its caption, a `Form`'s is the window
/// title, a `GroupBox`'s is its frame caption — and a `Panel`, a `TreeView` or
/// a toolbar simply never shows one. Getting this wrong is not cosmetic: the
/// property still round-trips (it lands on `data-text`), but PAINTING it puts
/// the control's own name inside it and, for a composed control, replaces the
/// chrome it was built with.
fn text_is_unpainted(class_name: &str) -> bool {
    matches!(
        class_name.to_ascii_lowercase().as_str(),
        "panel"
            | "flowlayoutpanel"
            | "tablelayoutpanel"
            | "splitcontainer"
            | "splitter"
            | "tabcontrol"
            | "treeview"
            | "listview"
            | "listbox"
            | "checkedlistbox"
            | "combobox"
            | "picturebox"
            | "progressbar"
            | "trackbar"
            | "hscrollbar"
            | "vscrollbar"
            | "monthcalendar"
            | "datagridview"
            | "datagrid"
            | "propertygrid"
            | "menustrip"
            | "toolstrip"
            | "statusstrip"
            | "contextmenustrip"
            | "bindingnavigator"
            | "webbrowser"
            | "usercontrol"
    )
}

fn accessor_node(
    target: &vybe_runtime::component_model::HostTarget,
    prop: &str,
    is_control: bool,
    class_name: &str,
) -> NamespaceNode {
    let setting = target.name == vybe_compiler::primitives::gui::HOST_FN_SET_PROPERTY;
    if is_control
        && (setting || target.name == vybe_compiler::primitives::gui::HOST_FN_GET_PROPERTY)
    {
        let role = match gui_property_role(prop) {
            "" => prop.to_ascii_lowercase(),
            // **Most controls INHERIT `Text` and never draw it.** A designer
            // writes one on everything it generates, so a container took its
            // own name as content: a `<div>` reading `pnl1`, a `<ul>` reading
            // `tvw1`, and — where the control is composed — a caption that
            // REPLACED its chrome, since `textContent` replaces all children
            // (DOM §4.4).
            //
            // Declared here because it is WinForms' own vocabulary: `Label`,
            // `Button`, `CheckBox`, `GroupBox` and `Form` paint their `Text`
            // and a `Panel` or `TreeView` does not. The shared role only has
            // to know that this one is not painted.
            //
            // ⚠ `ListBox`/`ComboBox`/`CheckedListBox` are here for a second
            // reason: their `Text` is the SELECTED item, and writing it as a
            // text child of a `<select>` is invalid markup that would sit
            // among the options.
            "text" if text_is_unpainted(class_name) => "unpaintedtext".to_string(),
            r => r.to_string(),
        };
        let prefix = if setting {
            vybe_compiler::primitives::gui::PROP_SET_EMIT
        } else {
            vybe_compiler::primitives::gui::PROP_GET_EMIT
        };
        NamespaceNode::CommonEmit(format!("{prefix}{role}"))
    } else {
        namespaces::host_fn(&target.module, &target.name)
    }
}

/// The HTML element a WinForms control IS — `tag`, or `tag:input-type` for the
/// `<input>` family. `None` means "not converted yet".
///
/// **Only controls with a genuine, conforming HTML counterpart are listed.**
/// A `Button` IS `<button>` and a `CheckBox` IS `<input type="checkbox">` —
/// same element plib maps `TButton`/`TCheckBox` to, which is what makes a
/// Delphi and a WinForms checkbox the same control rather than two lookalikes.
/// These get real form association, native keyboard semantics and a real
/// `value`, none of which a `<vybe-checkbox>` has.
///
/// Everything unlisted keeps falling through to `ControlElement::custom` →
/// `<vybe-listview>`, exactly as before. That is deliberate: a `ListView`,
/// `TreeView` or `DataGridView` has no HTML counterpart, and forcing one onto
/// `<table>` or `<div>` would claim semantics it does not have. Converting one
/// class at a time is what keeps this measurable — there is no flag day.
///
/// Kept in step with `platforms/plib/src/emitter/gcl/mod.rs`: a control that
/// both frameworks name must land on the SAME element, or the two frontends
/// quietly diverge on what is supposed to be one control.
fn html_element_for_control(class_name: &str) -> Option<&'static str> {
    Some(match class_name.to_ascii_lowercase().as_str() {
        // ── Real HTML, same elements plib maps the VCL twins to ────────────
        // A form IS the document's body, not a child element.
        // `emit_control_element` special-cases this: `createElement("body")`
        // is legal but yields a DETACHED second body that renders nothing, so
        // it takes `document.body` instead.
        "form" => "body",
        "button" => "button",
        "checkbox" => "input:checkbox",
        "radiobutton" => "input:radio",
        "textbox" => "input:text",
        "maskedtextbox" => "input:text",
        "richtextbox" => "textarea",
        "label" => "label",
        // A LinkLabel IS a hyperlink.
        "linklabel" => "a",
        "groupbox" => "fieldset",
        // The strips. `menu` is a real HTML element the document already knows
        // (`control_kind` maps it to the `menustrip` widget, born `Dock::Top`);
        // the `vybe-*` pseudo-tags these used to fall through to matched NO
        // `control_kind` arm, so every WinForms menu came out a 120x20 LABEL
        // stacked at the origin. plib spells `TMainMenu`/`TPopupMenu`/
        // `TMenuItem` the same way, and a ToolStrip is what `<menu>` is
        // specified to be — "a toolbar" — with WinForms' own `Dock=Top` default.
        // An ITEM is the same `menu` tag as the strip, on purpose — plib says
        // the same thing about `TMenuItem`. A bar word and the submenu it opens
        // are one element here, and the strip derives its words from its
        // children's captions at paint time.
        "menustrip"
        | "toolstrip"
        | "toolstripmenuitem"
        | "toolstripdropdownitem"
        | "toolstripbutton"
        | "toolstriplabel"
        | "toolstripstatuslabel" => "menu",
        // A separator IS a horizontal rule. Held on `menu` until `hr` had a
        // `control_kind` arm, because a tag with no arm is a silent 120x20
        // label — it has one now (`panel`, 200x2), so this is a rule.
        "toolstripseparator" => "hr",
        // A StatusStrip is the document's footer. `<footer>` has a
        // `control_kind` arm (a container) where `vybe-statusstrip` has none;
        // the `statusstrip` widget itself is still unreachable because no tag
        // maps to it, which is `widgets`' half of the mapping, not this
        // registrar's.
        "statusstrip" => "footer",
        "combobox" => "select",
        // **A ListBox and a ComboBox are ONE element, one attribute apart.**
        // HTML §4.10.7: a `<select>` with `size` above one is a list box,
        // without it a dropdown. This was a `<ul>` — which renders items and
        // has no selection model whatever, so a list could be looked at and not
        // used, while the ComboBox beside it worked. `size` is the whole
        // difference, and `SelectedIndex`/`SelectedItem`/`SelectedIndexChanged`
        // now resolve against the same `<select>` surface the ComboBox uses.
        "listbox" => "select;@size=4",
        // A tree is a nested list, and an EMPTY one is still a control: the
        // border and white field are what the user sees before a single node
        // is added, and what a bare `<ul>` does not draw.
        "treeview" => "ul;border:1px solid #c8c8c8;background-color:#ffffff;overflow:auto",
        // HTML has these outright, and they carry real semantics a `<div>`
        // cannot: a range input is keyboard-operable and `<progress>` is
        // announced as a progress indicator.
        "progressbar" => "progress",
        "trackbar" => "input:range",
        "numericupdown" => "input:number",
        // A PictureBox IS a drawing surface, and HTML spells that `<canvas>`.
        // `classes/media.rs` says the same thing from the other side: that
        // surface "is exactly what the `widgets::Canvas` widget provides".
        //
        // Only a `<canvas>` node owns a recording — `Document::canvas_mut`
        // gates on `control_kind == "canvas"` — so a control that is any other
        // tag has nothing for `getContext` to bind to and every `Graphics` call
        // is dropped by the backend, silently. A custom element cannot stand in:
        // the drawing surface is what the element IS, not a property it carries.
        //
        // ⚠ This is the DRAWING half only. `pb.Image = …` has no entry in
        // `gui_property_role`, so it lands on an attribute named `image` —
        // where an unmapped property belongs, and not something any element
        // renders. Displaying an assigned image means giving `image` a role
        // that reaches `drawImage` on this surface.
        //
        // ⚠ `<canvas>` has a UA default size of 300x150 where a custom element
        // takes the generic control default, so a PictureBox that never sets
        // `Size` changes size. Designer-generated forms always set it.
        "picturebox" => "canvas",

        // ── No HTML counterpart: a DECLARED custom element ─────────────────
        // Declaring `vybe-*` explicitly is better than letting it fall through
        // to `ControlElement::custom`, because the choice becomes visible here
        // instead of being implied by absence — and it keeps these names in
        // step with plib, which spells the same controls the same way
        // (`TImage`, `TSplitter`, `TPageControl`, `TTabSheet`, `TTimer`).
        // The scrollbars and the navigator. `widgets` has had all three
        // kinds and their default sizes the whole time; what was missing was
        // the DECLARATION, without which they had no `CtorSpec` and every
        // property write on them was silently dropped.
        // ⚠ A ContextMenuStrip is NOT `<menu>`, and the difference is docking.
        // The `menu` TAG is born `Dock::Top`, which is right for a menu bar and
        // wrong for a popup: `cms1` took the full width under the other strips
        // and threw away the `Location`/`Size` the designer gave it, pushing
        // the first real control off the top of the form. A context menu is
        // attached to a control and shown on demand — it is not a bar.
        // A context menu IS a `<menu>` — what differs from the bar is that it
        // is not shown until something opens it, and `display:none` is how HTML
        // says that. The custom tag was standing in for the DOCKING difference
        // described below, which a declaration can state directly.
        "contextmenustrip" => "menu;display:none;position:absolute",
        // Declared `vybe-*` custom elements. `control_kind` strips `vybe-` and
        // looks the remainder up against the widget list, so the TAG carries
        // the kind and these two land on real widgets that already exist
        // (`checkedlistbox`; `datagrid` folds onto `datagridview`).
        // The same list, multi-select: HTML's own way to spell "several of
        // these at once", and a strict improvement on the `<ul>` this was —
        // items are shown AND selectable instead of only shown.
        //
        // ⚠ Not the whole control. A per-item CHECKBOX is the one thing
        // `<select>` does not have, and `CheckedIndices`/`CheckedItems` are
        // declared as property names with nothing registered behind them, so
        // the checked API answers nothing on either mapping. Wiring it means
        // deciding whether "checked" reads the selection or the control grows
        // a `<ul>` of `<input type=checkbox>` — an open question, not something
        // this mapping settles.
        "checkedlistbox" => "select;@size=4;@multiple",
        // The legacy grid is the same control and takes the same element, or
        // the two spellings would render differently for no reason.
        "datagrid" => "table;border-collapse:collapse;border:1px solid #c8c8c8;background-color:#ffffff",
        // ⚠ These have no widget kind YET, so they render as a label until
        // `widgets` grows one — the designed degradation, visible in a
        // capture instead of the control vanishing. The declaration still buys
        // construction, identity, geometry, text, events and data binding.
        // A property grid is a two-column table of name/value rows, which is
        // what a `<table>` is. Its rows arrive at runtime like a grid's.
        "propertygrid" => "table;border-collapse:collapse;background-color:#ffffff",
        // The bare drag bar (plib's `TSplitter`), not the two-panel container.
        "splitter" => "div;background-color:#c8c8c8;cursor:col-resize",
        // A text field you step through a list of strings with — the same
        // control as `NumericUpDown` over words instead of numbers.
        "domainupdown" => "input:text",
        // A UserControl is a plain composite container, and `<section>` is a
        // real element that already establishes a containing block.
        "usercontrol" => "section",
        // **A scrollbar is a value in a range**, which is what `<input
        // type=range>` IS — `Minimum`/`Maximum`/`Value`/`LargeChange` and the
        // `Scroll` event map onto it one for one, the same way `TrackBar`
        // already does. The vertical one says so in CSS rather than in a
        // different tag: one control, one element, an axis.
        "hscrollbar" => "input:range",
        "vscrollbar" => "input:range;writing-mode:vertical-lr",
        // A toolbar of buttons around a position box — the markup is declared
        // in `default_markup_for_control`, which is what makes it a navigator
        // rather than an empty strip.
        "bindingnavigator" => "menu;display:flex;align-items:center;gap:2px",
        // ⚠ `vybe-splitter`, not `vybe-splitcontainer`, made this a LABEL.
        // The tag carries the kind: `control_kind` strips `vybe-` and looks the
        // remainder up against the widget list, which spells it
        // `splitcontainer`. A tag naming no known control degrades to a 120x20
        // label — so a mapping that renames the control silently deletes it.
        // plib's `TSplitter` is a different control (a bare drag bar); this one
        // is WinForms' two-panel container.
        // Two panes and a splitter, declared as markup — a flex row whose
        // children ARE `Panel1` and `Panel2`.
        "splitcontainer" => "div;display:flex;flex-direction:row",
        // A tab strip over a page. Both are ordinary containers; which page
        // shows is a `display` question the control answers at runtime.
        "tabcontrol" => "div;display:flex;flex-direction:column;border:1px solid #c8c8c8;background-color:#ffffff",
        "tabpage" => "section",
        // A Timer is a `components` member like the providers below: present and
        // scriptable, never painted.
        "timer" => "div;display:none",

        // ── The layout panels ──────────────────────────────────────────────
        // **Neither is a control — both are a `<div>` and a `display` mode.**
        // A `Panel`, a `FlowLayoutPanel` and a `TableLayoutPanel` hold children
        // and draw a background; what separates them is how those children are
        // arranged, and CSS has said that since flexbox and grid. So they are
        // the element `Panel` already is, plus the CSS that makes them differ.
        //
        // This is the browser-swap test in one line. A `<vybe-tablelayoutpanel>`
        // — which is what these were — needs a `customElements.define` and a
        // layout implementation before a real engine renders it at all; a
        // `<div style="display: grid">` needs neither, because every browser
        // already implements it. The custom tag was the workaround `declares`
        // exists to remove: `ControlElement` carries declared CSS through
        // `setStyleProperty`, so it cascades and serialises like any author
        // style.
        //
        // ⚠ A `TableLayoutPanel` is a GRID, not a `<table>`. Its cells are
        // positions in a fixed set of tracks, not content that sizes them —
        // `ColumnCount` says how many, where a table's columns come from what
        // its cells hold. Mapping it to `<table>` would make an empty grid.
        //
        // `ColumnCount`/`RowCount` already reach `grid-template-columns`/
        // `-rows` as `repeat(n, 1fr)` (the shared role table in
        // `primitives/gui.rs`), so an explicit count arrives on its own and
        // OVERWRITES the declared default below, exactly as a later CSS
        // declaration would.
        //
        // ⚠ The 2x2 default is not decoration. These used to be constructed as
        // `TableLayoutPanel::new(2, 2)`, so a designer that never writes
        // `ColumnCount` — `examples/vb/allcontrols` is one — got two columns
        // from the constructor. A bare `display: grid` would have collapsed it
        // to a single column and lost the layout silently, which is the exact
        // failure the old "deliberately NOT converted" comment feared. The
        // default moves from a constructor argument to a declaration; it does
        // not disappear.
        //
        // The values match `widgets/src/html/panel.rs` (`table_css`,
        // `flow_css`), which has spelled the same CSS for these two controls
        // all along — one fact, not two.
        "flowlayoutpanel" => "div;display:flex;flex-wrap:wrap;gap:4px",
        "tablelayoutpanel" => {
            "div;display:grid;grid-template-columns:repeat(2, 1fr);grid-template-rows:repeat(2, 1fr);gap:2px"
        }

        // ── Controls with a real HTML counterpart ──────────────────────────
        // Native elements, not `vybe-*`, because HTML already has them and
        // `control_kind` already routes them to the right widget: an `<input
        // type=datetime-local>` IS the datetimepicker. Inventing a custom tag
        // for it would be a second spelling for something the platform ships.
        "datetimepicker" => "input:datetime-local",

        // **A `DataGridView` is NOT a `<table>`.** This used to emit `table`,
        // on the reasoning that a table "IS" the grid. It is the reverse
        // conflation of the one `control_kind` carried on the other side: a
        // `<table>` is a LAYOUT — a box that arranges its rows and cells —
        // whereas a `DataGridView` is a CONTROL, with columns it defines
        // itself, a header it draws itself and scrolling of its own. Now that
        // `<table>` means the layout, emitting `table` here would render every
        // WinForms grid as an empty table box with no columns at all.
        //
        // **A `DataGridView` IS a `<table>`** — now that a `<table>` is one.
        //
        // This is the second half of the same correction. `<table>` used to
        // mean the datagridview WIDGET, which made an HTML layout impersonate
        // a .NET control; the widget layer now implements CSS tables, so the
        // honest mapping runs the other way and the control is the markup.
        // Its columns and rows are `<th>`s and `<tr>`s, built by
        // `datagrid_adapter` from `Columns.Add`/`Rows.Add`.
        //
        // ⚠ It went through `vybe-datagridview` on the way here, which was
        // correct while the data surface was inert but is not a destination: a
        // `vybe-*` tag needs a `customElements.define` AND a grid
        // implementation before a real engine draws anything, where a `<table>`
        // needs neither. The custom tag rendered the grid's default chrome —
        // a header band and empty row lines — and no data could reach it,
        // because a widget does not render DOM children.
        //
        // `border-collapse` is the grid look: one line between cells rather
        // than two, which is what a data grid draws.
        "datagridview" => "table;border-collapse:collapse;border:1px solid #c8c8c8;background-color:#ffffff",
        // A ListView in its Details mode is a table with a header row, which is
        // what `DataGridView` above already resolves to and what its `Columns`
        // and `Items` append into. The other view modes are a `display`
        // difference over the same items, not a different control.
        "listview" => "table;border-collapse:collapse;border:1px solid #c8c8c8;background-color:#ffffff",
        // A month grid — the chrome is declared in `default_markup_for_control`.
        "monthcalendar" => "div;border:1px solid #c8c8c8;background-color:#ffffff",

        // ── Non-visual components and the dialogs ──────────────────────────
        // A Timer, ToolTip or file dialog is a member of the form, not a box
        // on it — WinForms puts them in `components`, not `Controls`, and
        // neither ever paints. `renders_nothing` in `widgets` already
        // names every tag below, so they are nodes that occupy no rectangle:
        // constructible, nameable, event-wireable, and invisible.
        //
        // That is the whole reason they can come off the factory. The element
        // was never the problem — a control that drew a grey label where a
        // timer should be was, and the widget side fixed that before this.
        // A `WebBrowser` IS an `<iframe>` — an embedded browsing context, which
        // is what the control is and what a real engine already implements. No
        // `vybe-*` tag: inventing one for something HTML has outright is the
        // hack this conversion exists to remove.
        //
        // ⚠ `control_kind` maps `iframe` to the `picturebox` widget, because
        // `widgets` has no `webbrowser` kind. So it renders as a plain
        // box until one exists — Youness's call, deliberately taken: the tag is
        // right, the widget is a later job. When that widget lands, the only
        // change needed is a `control_kind` arm; this mapping stays.
        "webbrowser" => "iframe",
        // **"Renders nothing" is a CSS answer, not a tag we had to invent.**
        // `display: none` is HTML's own way to say a node is present, scriptable
        // and unpainted — exactly what a `components` member is — and it needs
        // no `customElements.define` for a real engine to honour it. These were
        // `vybe-*` only because the widget side kept a `renders_nothing` list of
        // tag names; a declaration says the same thing in the cascade, where a
        // browser already reads it.
        //
        // An ImageList holds `<img>` children that are never shown, which is
        // what a hidden container IS.
        "imagelist" | "tooltip" | "notifyicon" | "errorprovider" | "helpprovider"
        | "backgroundworker" => "div;display:none",
        // **The file dialogs ARE `<input type=file>`.** A web page opens a file
        // chooser by clicking a hidden file input — that is not a polyfill, it
        // is how the platform exposes the OS chooser, and `ShowDialog` is that
        // click. The control keeps its `FileName`/`Filter` surface; what it
        // stops needing is a custom element standing in for a picker HTML has.
        // (`control_kind` already maps `file` to the `fileinput` widget, and
        // webcore draws the button-plus-label.)
        "openfiledialog" | "savefiledialog" | "folderbrowserdialog" => "input:file;display:none",
        // Likewise the colour chooser: `<input type=color>` IS one, with the
        // swatch and the picker webcore already implements.
        "colordialog" => "input:color;display:none",
        // HTML has no font picker, but it does have the element a dialog IS.
        // A `<dialog>` without `open` is not rendered (HTML §4.11.4), so this
        // is invisible until something shows it — the control's own semantics,
        // spelled by the element rather than by a list of tag names.
        "fontdialog" => "dialog",

        // A `Panel` IS a `<div>` — a block container that draws a background
        // and holds children, which is the whole of what the control is. plib
        // has mapped `TPanel` this way all along; the two now agree.
        //
        // It was held back because `<div>` containers laid out wrong in the
        // shared engine: every container ran the flexbox algorithm regardless
        // of `display`, so a `<div>` behaved as `display: flex; flex-direction:
        // column; align-items: stretch` and a row of children came out as a
        // column of full-width bars. `widgets` grew real CSS normal flow
        // on 2026-08-15 (`Formatting::{Flex,Normal}` on `FlowLayoutPanel`, set
        // from the computed `display`), which is the defect that gate was
        // waiting on.
        "panel" => "div",

        // ── Everything else ────────────────────────────────────────────────
        // A class with no element here is not a control this platform can
        // construct as a node, and answering None is what says so.
        //
        // ⚠ This used to hold `FlowLayoutPanel` and `TableLayoutPanel` back,
        // on the reasoning that a `display` mode "needs a selector" and the
        // control's role was a struct field no selector could match. That was
        // true of a SHEET and never of the element: `ControlElement.declares`
        // carries CSS the control is born with, which needs no selector at
        // all. Both are converted above.
        _ => return None,
    })
}

/// The declared parent of a descriptor class — the ONE place this registrar
/// reads the inheritance link.
///
/// Every derivation below (identity chain, flattened methods, flattened
/// properties, "is this a control") is the same self-first walk over this
/// link, so they share `namespaces::ancestry_of` / `namespaces::inherits`
/// instead of each re-writing the walk with its own cycle guard. There were
/// four copies of it in this file.
fn descriptor_parent(class_name: &str) -> Option<String> {
    crate::emitter::surface()
        .component_descriptor()
        .classes
        .iter()
        .find(|c| c.name.eq_ignore_ascii_case(class_name))
        .and_then(|c| c.parent.clone())
}

/// The identity chain for a control, self first — `["Button", "ButtonBase",
/// "Control", …]`. Stamped as `__type` (first) + `__types` (all), so `is` /
/// `instanceof` answer for every ancestor from the shared reflection path
/// rather than a per-language fallback.
fn control_ancestry(class_name: &str) -> Vec<String> {
    namespaces::ancestry_of(class_name, descriptor_parent)
}

/// The generic-construction spec for a converted control — the same shape plib
/// and flutter register, which is what lets `registered_control_element` answer
/// for dotnet at last.
///
/// `params`/`fields` are deliberately EMPTY: a WinForms control is constructed
/// bare (`new Button()`) and configured by property assignment afterwards,
/// unlike a Flutter widget whose children arrive as constructor arguments. So
/// `control_fn` is the only thing making this a construction, and the object
/// IS the element.
fn control_ctor_spec(class_name: &str, element: &str) -> vybe_runtime::namespaces::CtorSpec {
    vybe_runtime::namespaces::CtorSpec {
        params: Vec::new(),
        fields: Vec::new(),
        field_gui: Vec::new(),
        ancestry: control_ancestry(class_name),
        control_fn: Some(element.to_string()),
        inner_html: default_markup_for_control(class_name).map(str::to_string),
        // A WinForms control IS its element at construction; nothing to inflate.
        nest_coerce: None,
        value_equality: false,
    }
}

/// The chrome a composite control is BORN with, as HTML.
///
/// **.NET's own constructor builds these, which is why a designer file never
/// does.** `BindingNavigator`'s ctor runs `AddStandardItems`; `SplitContainer`
/// comes with `Panel1`/`Panel2` already there — `sc.Panel1.Controls.Add(x)`
/// only works because the pane EXISTS; a `MonthCalendar` is a month grid the
/// moment you make one. A declaration that could name only a tag left each of
/// them one empty box, which is exactly how they rendered: a bare label with
/// the control's name and nothing inside it.
///
/// Written as markup because the chrome IS static HTML. Every part is a real
/// element — selectable, stylable, and able to take a listener — so wiring a
/// navigator button to its `BindingSource` is `querySelector` + a listener,
/// the same two calls a script would make. Children built from RUNTIME values
/// stay with their adapter instead (`DataGridView`'s `Columns.Add`).
///
/// ⛔ No `id` anywhere: an id must be unique per document (DOM §4.9), and this
/// template is instantiated once per control, so ids would collide the moment
/// a form held two navigators and quietly break `getElementById` for both.
/// The parts carry CLASSES, which are as addressable and stay valid.
///
/// The styling is the CONTROL's appearance, not HTML's — the same argument
/// `datagrid_adapter` makes about its cell borders. It rides in the markup's
/// own `style` attributes rather than the UA sheet, so it cascades normally
/// and an author's later write simply overrides it.
/// A `ToolStripButton`'s resting appearance: a glyph on the strip, with the
/// strip's own background showing through. WinForms only draws a border and a
/// fill on hover or press, so the resting state has neither.
///
/// A macro rather than a `const`, because `concat!` takes literals only and
/// the markup below is built with it.
macro_rules! toolstrip_button {
    () => {
        "min-width:23px;height:22px;padding:0 4px;border:none;\
         border-radius:0;background-color:transparent"
    };
}

fn default_markup_for_control(class_name: &str) -> Option<&'static str> {
    Some(match class_name.to_ascii_lowercase().as_str() {
        // The standard items, in WinForms' order: move-first, move-previous,
        // the position box reading "N of M", move-next, move-last.
        //
        // The items are `ToolStripButton`s, and a ToolStripButton is FLAT:
        // no border and no fill until the pointer is over it. Left with the
        // UA's `<button>` chrome they drew as rounded grey lozenges — right
        // element, wrong control. `TOOLSTRIP_BUTTON` below is that appearance,
        // and it has to name every property the UA sheet sets, because a
        // shorthand it does not mention keeps the UA's value: `border-radius`
        // is the one that made them lozenges and `border`/`background` alone
        // would have left it.
        "bindingnavigator" => concat!(
            "<button type='button' class='vybe-nav vybe-nav-first' style='", toolstrip_button!(), "'>|&#9664;</button>",
            "<button type='button' class='vybe-nav vybe-nav-prev' style='", toolstrip_button!(), "'>&#9664;</button>",
            "<input type='text' class='vybe-nav-position' value='0 of 0'",
            " style='width:64px;text-align:center;margin:0 4px;height:21px;border-radius:0",
            ";border:1px solid #7a7a7a;background-color:#ffffff'>",
            "<button type='button' class='vybe-nav vybe-nav-next' style='", toolstrip_button!(), "'>&#9654;</button>",
            "<button type='button' class='vybe-nav vybe-nav-last' style='", toolstrip_button!(), "'>&#9654;|</button>"
        ),
        // Two panes and the splitter between them. `Panel1`/`Panel2` resolve to
        // these, so they must exist before any code adds a control to one.
        "splitcontainer" => concat!(
            "<div class='vybe-split-panel1' style='flex:1 1 50%;overflow:auto'></div>",
            "<div class='vybe-splitter' style='flex:0 0 4px;background-color:#c8c8c8;cursor:col-resize'></div>",
            "<div class='vybe-split-panel2' style='flex:1 1 50%;overflow:auto'></div>"
        ),
        // ⛔ No `monthcalendar` arm. A calendar has no static spelling: the
        // month it opens on is the date the program RUNS. That makes it
        // BEHAVIOUR, and behaviour belongs to `platforms/web`, which owns the
        // relationship with the browser — this crate declares only what the
        // control IS.
        _ => return None,
    })
}

/// Does this class DESCEND FROM `Control` — i.e. is it a control at all?
///
/// The element model applies to controls and nothing else. `Point`, `Size`,
/// `Font`, `Pen`, `Brush` and `Graphics` are value types and a drawing surface:
/// they are not elements and have no DOM counterpart. Their members compose in
/// bytecode (`dotnet.point_new`, `dotnet.pen_new`) and their properties are
/// ordinary struct fields.
///
/// Getting this wrong is not subtle. Routing them through the roles turned
/// `Size.Width` into a CSS geometry write and `Point.X` into an attribute, so
/// `new Point(100, 200).x` stopped reading back — caught by
/// `winforms/new_point_properties` and `new_point_and_size`, which is exactly
/// what those tests are for.
fn descends_from_control(class_name: &str) -> bool {
    namespaces::inherits(class_name, "Control", descriptor_parent)
}

/// Is this class a NODE IN THE DOCUMENT — the question the roles actually ask.
///
/// `descends_from_control` answers a .NET inheritance question, and for almost
/// every class the two coincide. `ToolStripItem` is where they part: .NET
/// derives it from `Component`, not `Control`, so a menu item is not a control
/// — but it is unquestionably an element, and its `Text` is a caption on that
/// element like every other caption.
///
/// Declaring it a `Control` to get the roles would have been the cheap fix and
/// the wrong one: `mi is Control` must answer False, because that is what .NET
/// answers. So the hierarchy stays truthful and the GATE moves to the property
/// that decides it — being element-mapped.
///
/// This does NOT loosen the value-type exclusion documented above.
/// `Point`/`Size`/`Font`/`Pen`/`Brush`/`Graphics` have no arm in
/// [`html_element_for_control`], so they answer `false` from both halves.
fn element_backed_control(class_name: &str) -> bool {
    descends_from_control(class_name) || html_element_for_control(class_name).is_some()
}

/// Does this class map to an element? The descriptor builder asks, so an
/// element-mapped class gets a constructor with no host factory behind it —
/// see `winforms/component_classes.rs`.
pub(crate) fn is_element_mapped(class_name: &str) -> bool {
    html_element_for_control(class_name).is_some()
}

/// A class's own methods followed by every inherited one, nearest first.
///
/// Registration-time expansion is the tree's DECLARED contract, not a
/// workaround for a missing edge: `lookup_type_instance_member` is a flat map
/// get by design, and `plib`'s registrar states the rule — "the adapter knows
/// its own inheritance, nothing downstream should". A namespace tree resolves
/// PATHS; class inheritance is a different relation, and the adapter owns it
/// because only the adapter has the descriptor's `parent` chain.
///
/// So a class's node must carry its whole inherited surface or the member is
/// simply unreachable. `inherited_properties`
/// below already did this for properties; methods were left flat, and that is
/// why `Button.Hide()` resolved to nothing: `Show`/`Hide`/`Focus` are declared
/// on `Control`, so `Button`'s node had no `hide` for
/// `lookup_type_instance_member` to find, the shared resolver answered `None`,
/// and the call fell back to a name-keyed `struct.get "Hide"` expecting a
/// ctor-bound thunk that the DOM construction path no longer builds.
///
/// `StringBuilder.Append()` worked throughout precisely because `Append` is
/// declared on `StringBuilder` itself — the same resolver, reached because the
/// member happened to be flat.
///
/// Nearest declaration first, so the `or_insert` folds at the call site give
/// override-shadows-base — the same rule real .NET virtual dispatch uses.
fn inherited_methods(class_name: &str) -> Vec<vybe_runtime::component_model::MethodDef> {
    declared_up_the_chain(class_name, |class| class.methods.clone())
}

/// A class's own properties followed by every inherited one, nearest first, so
/// an `or_insert` fold gives override-shadows-base.
fn inherited_properties(class_name: &str) -> Vec<vybe_runtime::component_model::PropertyDef> {
    declared_up_the_chain(class_name, |class| class.properties.clone())
}

/// Everything `select` declares on `class_name` and then on each ancestor, in
/// nearest-first order.
///
/// The chain comes from the shared `ancestry_of` walk, so methods, properties
/// and any future member kind agree on ordering and on what a cyclic `parent`
/// does — they used to be separate loops that happened to match.
fn declared_up_the_chain<T, F>(class_name: &str, select: F) -> Vec<T>
where
    F: Fn(&vybe_runtime::component_model::ClassType) -> Vec<T>,
{
    let descriptor = crate::emitter::surface().component_descriptor();
    let mut out = Vec::new();
    for name in namespaces::ancestry_of(class_name, descriptor_parent) {
        if let Some(class) = descriptor
            .classes
            .iter()
            .find(|c| c.name.eq_ignore_ascii_case(&name))
        {
            out.extend(select(class));
        }
    }
    out
}

/// Properties whose accessors are shared EMITS rather than host calls. They
/// have no descriptor `PropertyDef` — they were a hand-written cascade in the
/// dotnet surface — so they register here as the data they always were.
fn shared_emit_accessors(class_name: &str) -> Vec<(String, NamespaceNode)> {
    let emit = |name: &str| NamespaceNode::CommonEmit(name.to_string());
    let rw = |g: &str, s: &str| namespaces::property(Some(emit(g)), Some(emit(s)));
    let ro = |g: &str| namespaces::property(Some(emit(g)), None);
    let entries: &[(&str, NamespaceNode)] = &match class_name.to_ascii_lowercase().as_str() {
        "stringbuilder" => vec![
            // The INDEXER, as the data it always was. `sb[i]` and `sb[i] = c`
            // were a hand-written pair in the shared compiler, selected by a
            // language-family check; declaring both directions here lets the
            // index sites ask the tree which accessors the receiver's type
            // owns. Both directions are declared together on purpose — a read
            // and a write that disagree about the storage shape is the bug the
            // pair was gated to avoid.
            ("Item", rw("dotnet.sb_index_get", "dotnet.sb_index_set")),
            ("Length", rw("dotnet.sb_length", "dotnet.sb_set_length")),
            ("Capacity", rw("dotnet.sb_capacity", "dotnet.sb_set_capacity")),
            ("MaxCapacity", ro("dotnet.sb_max_capacity")),
        ],
        // `Lazy(Of T)` — READ-ONLY computed properties, which is what makes
        // `.Value` run the factory on first read. An eager struct field could
        // not keep `IsValueCreated` False until someone asks.
        "lazy" => vec![
            ("Value", ro("dotnet.lazy_value")),
            ("IsValueCreated", ro("dotnet.lazy_is_value_created")),
        ],
        "stopwatch" => vec![
            ("ElapsedMilliseconds", ro("dotnet.stopwatch_elapsed_ms")),
            ("ElapsedTicks", ro("dotnet.stopwatch_elapsed_ticks")),
            ("Elapsed", ro("dotnet.stopwatch_elapsed")),
            ("IsRunning", ro("dotnet.stopwatch_is_running")),
        ],
        "task" => vec![
            ("Result", ro("dotnet.task_result")),
            ("IsCompleted", ro("dotnet.task_is_completed")),
            ("IsCanceled", ro("dotnet.task_is_canceled")),
            // `TaskStatus`, as the .NET spelling of the promise's own state.
            ("Status", ro("dotnet.task_status")),
            ("IsFaulted", ro("dotnet.task_is_faulted")),
        ],
        // `TaskCompletionSource(Of T).Task` — the consumer half. READ-ONLY and
        // computed: a struct field named `Task` on the source would be read as
        // `task` by a case-insensitive front end, and the accessor keeps the
        // promise reachable under the declared spelling for C# too.
        "taskcompletionsource" => vec![("Task", ro("dotnet.tcs_task"))],
        "weakreference" => vec![
            (
                "Target",
                rw("dotnet.weakref_target", "dotnet.weakref_set_target"),
            ),
            ("IsAlive", ro("dotnet.weakref_is_alive")),
        ],
        "list" | "arraylist" => vec![("Capacity", ro("dotnet.list_capacity"))],
        // `MemoryStream` — every one of these is DERIVED. `Length` changes on
        // each write, `Position` is the cursor, and `Capacity`'s setter has to
        // resize the backing store and can refuse.
        "memorystream" => vec![
            ("Capacity", rw("dotnet.ms_capacity", "dotnet.ms_set_capacity")),
            ("Length", ro("dotnet.ms_length")),
            ("Position", rw("dotnet.ms_position", "dotnet.ms_set_position")),
            ("CanRead", ro("dotnet.ms_can_read")),
            ("CanWrite", ro("dotnet.ms_can_write")),
            ("CanSeek", ro("dotnet.ms_can_seek")),
        ],
        // The two members a cursor cannot store — both are derived from
        // whatever `DataSource` currently points at, so a field would go stale
        // the moment the source changed. `Position`, `DataMember`, `Filter` and
        // `Sort` are NOT here: those are real fields the constructor writes.
        "bindingsource" => vec![
            ("Count", ro("dotnet.bindingsource_count")),
            ("Current", ro("dotnet.bindingsource_current")),
        ],
        // ── Strips and their items ─────────────────────────────────────────
        // `Items` IS the strip: WinForms wraps the contents in a
        // `ToolStripItemCollection`, but the `<menu>` element already is that
        // container, so the getter hands back the receiver and allocates
        // nothing. What makes the NEXT hop work is the declared return type
        // (see `self_member_returns`) — `ms.Items.Add(x)` looks `Add` up on
        // `ToolStripMenuItem`, and an item and a strip are the same element.
        //
        // `Add` is a CHILD append, not an item append: a strip is handed a
        // control the caller built, where a `ListBox` is handed text and makes
        // the option itself. The call site cannot tell them apart, which is why
        // each class declares which one it means rather than the emit guessing
        // from the argument.
        //
        // Arity travels with the node — a bare `CommonEmit` leaf is found as a
        // NAME and then not called, which is how `b.Hide()` once reached
        // "undefined is not callable".
        "menustrip"
        | "toolstrip"
        | "statusstrip"
        | "contextmenustrip"
        | "toolstripmenuitem"
        | "toolstripdropdownitem" => vec![
            ("Items", ro("dotnet.self")),
            ("DropDownItems", ro("dotnet.self")),
            // ⚠ The DECLARED .NET spelling. These read `"add"` / `"capacity"`
            // while every sibling here reads `Items` / `Length` / `MaxCapacity`,
            // which is the fold the tree used to require: a lowercase key was
            // the only kind a `get(&name.to_lowercase())` could ever find.
            // Exact-first lookup made that a liability rather than a rule — C#
            // writes `Add` and now MISSES the exact key, resolving only because
            // `fold_get` tries the fold afterwards. Correct spelling means C#
            // hits directly and VB still folds to it (casesensitivityplan §5d).
            (
                "Add",
                namespaces::overloads(vec![(
                    2,
                    emit(vybe_compiler::primitives::gui::APPEND_CHILD_EMIT),
                )]),
            ),
        ],
        // ── The grid's data surface ────────────────────────────────────────
        // `Columns` and `Rows` hand back the GRID, the same `dotnet.self`
        // shape the strips use for `Items` — a collection object would be a
        // second place for state the document already holds. What separates
        // them is the declared return TYPE (`self_member_returns`), because
        // unlike a strip's items these two do NOT mean the same append:
        // `Columns.Add` builds a `<th>`, `Rows.Add` a whole `<tr>` of `<td>`s.
        // Aliasing them to one `Add` the way the strips do would make a column
        // and a row the same thing, which is exactly what they are not.
        "datagridview" | "datagrid" => vec![
            ("Columns", ro("dotnet.self")),
            ("Rows", ro("dotnet.self")),
        ],
        // The two collection types the members above read back as. They hold
        // nothing and are never constructed — they exist so that `Add` has
        // somewhere to be found, and so that it can be a DIFFERENT `Add` on
        // each. Arity counts the receiver.
        "datagridviewcolumncollection" => vec![(
            "Add",
            namespaces::overloads(vec![
                (2, emit("dotnet.datagrid_add_column")),
                (3, emit("dotnet.datagrid_add_column")),
            ]),
        )],
        // Registered per arity because `Rows.Add(…)` is variadic and bytecode
        // has no variadic shape: each count is its own overload, and the emit
        // reads `argc` to know how many cells to build. Eight is a stated
        // ceiling — beyond it the call is not found, which is the loud answer.
        "datagridviewrowcollection" => vec![(
            "Add",
            namespaces::overloads(
                (1..=9)
                    .map(|argc| (argc, emit("dotnet.datagrid_add_row")))
                    .collect(),
            ),
        )],
        _ => vec![],
    };
    let mut out: Vec<(String, NamespaceNode)> = entries
        .iter()
        .map(|(n, node)| ((*n).to_string(), node.clone()))
        .collect();

    // ── `Controls` is the ELEMENT, `Add` is `appendChild` ──────────────────
    //
    // A control's children live in the DOM, so `Controls` allocates nothing and
    // hands back the receiver — the same `dotnet.self` shape the strips use for
    // `Items`. `self_member_returns` (at the registration site) declares that it
    // reads back as `Control`, which is where `Add` below is found.
    //
    // Declared here because the ONLY thing standing in for it was a hardcoded
    // `members[..] == "controls" && members[..] == "add"` in the shared call
    // path (`calls.rs`). That compares against canon'd names, and `canon` folds
    // to lowercase only for a profile with `case_sensitive = false` — so the
    // match fired for vb/pascal/cobol and was DEAD for csharp, whose canon
    // preserves `["Controls","Add"]`. Same source, same AST, opposite outcome:
    // `--dump` shows vb reach `appendChild` and csharp fall through to
    // `struct.get Controls` on the body element, which is undefined.
    //
    // The registered route has no such asymmetry — `lookup_type_instance_member`
    // folds the member before matching, so every language reaches it. A control
    // that is element-backed is exactly one that HAS DOM children, which is why
    // the gate is `element_backed_control` and not a name list.
    if element_backed_control(class_name) {
        let ro = |name: &str| {
            namespaces::property(Some(NamespaceNode::CommonEmit(name.to_string())), None)
        };
        // `Controls` and `Add`, in the spelling .NET declares — see the note
        // on the strips above.
        out.push(("Controls".to_string(), ro("dotnet.self")));
        // Arity counts the receiver, as every other node here does — a bare
        // `CommonEmit` leaf is found as a NAME and then not called, which is
        // the fault already recorded on `Hide`.
        out.push((
            "Add".to_string(),
            namespaces::overloads(vec![(
                2,
                NamespaceNode::CommonEmit(
                    vybe_compiler::primitives::gui::APPEND_CHILD_EMIT.to_string(),
                ),
            )]),
        ));
    }
    out
}

/// What a member that IS the receiver READS BACK as.
///
/// The load-bearing half of the `dotnet.self` members above: the getter hands
/// back the receiver, and this is what the NEXT hop resolves against. Without
/// it `ms.Items.Add(x)` resolves `Add` against nothing and calls `undefined`.
///
/// Declaring `ToolStripMenuItem` while the runtime holds the STRIP is sound
/// only because a strip and an item are the same `<menu>` element with the same
/// member set — the same deliberate alias plib documents on `TMainMenu.Items`.
/// Give an item its own tag and this stops being true.
fn self_member_returns(class_name: &str) -> &'static [(&'static str, &'static str)] {
    match class_name.to_ascii_lowercase().as_str() {
        "menustrip"
        | "toolstrip"
        | "statusstrip"
        | "contextmenustrip"
        | "toolstripmenuitem"
        | "toolstripdropdownitem" => &[
            ("Items", "ToolStripMenuItem"),
            ("DropDownItems", "ToolStripMenuItem"),
        ],
        // ⚠ Two DIFFERENT types, where the strips above alias one. A strip and
        // its item are the same `<menu>` element, so a single member set serves
        // both; a column and a row are not the same thing and their `Add`s
        // build different elements. Declaring one type for both would silently
        // make `Rows.Add` append a `<th>`.
        "datagridview" | "datagrid" => &[
            ("Columns", "DataGridViewColumnCollection"),
            ("Rows", "DataGridViewRowCollection"),
        ],
        // ── System.Collections.Immutable ──────────────────────────────────
        //
        // A persistent collection's whole surface CHAINS: every "mutation"
        // answers a new collection of the same type. Without these the value
        // coming back from `Add` is a bare array with no identity, so the
        // next hop resolves against nothing — `a1.Add(2).Length` died there
        // even with the type fully declared.
        "immutablearray" => &[
            ("Add", "ImmutableArray"),
            ("AddRange", "ImmutableArray"),
            ("RemoveAt", "ImmutableArray"),
            ("SetItem", "ImmutableArray"),
        ],
        "immutablelist" => &[
            ("Add", "ImmutableList"),
            ("AddRange", "ImmutableList"),
            ("RemoveAt", "ImmutableList"),
            ("SetItem", "ImmutableList"),
        ],
        // `Dequeue`/`Pop` answer the REMAINING collection, not the element —
        // .NET spells the element read as a separate `Peek`.
        "immutablequeue" => &[("Enqueue", "ImmutableQueue"), ("Dequeue", "ImmutableQueue")],
        "immutablestack" => &[("Push", "ImmutableStack"), ("Pop", "ImmutableStack")],
        "immutablehashset" => &[
            ("Add", "ImmutableHashSet"),
            ("Remove", "ImmutableHashSet"),
            ("Union", "ImmutableHashSet"),
            ("Intersect", "ImmutableHashSet"),
            ("Except", "ImmutableHashSet"),
        ],
        "immutabledictionary" => &[
            ("Add", "ImmutableDictionary"),
            ("SetItem", "ImmutableDictionary"),
        ],
        // The builder is the one type here that does NOT chain — its `Add`
        // mutates and returns nothing. Only the handoff back to immutability
        // carries a type.
        "immutablelistbuilder" => &[("ToImmutable", "ImmutableList")],
        _ => &[],
    }
}

/// The tree roots this platform registers under — what a `type_scopes`
/// consumer names to reach these classes.
#[cfg(test)]
fn dotnet_scope() -> Vec<String> {
    vec!["dotnet".to_string()]
}

fn console_stdout_writer_node() -> NamespaceNode {
    let mut statics = Subtree::new();
    statics.insert(
        "write".into(),
        NamespaceNode::CommonEmit("dotnet.console_write".into()),
    );
    statics.insert(
        "writeline".into(),
        NamespaceNode::CommonEmit("dotnet.console_writeline".into()),
    );
    NamespaceNode::Namespace(statics)
}

fn console_stderr_writer_node() -> NamespaceNode {
    let mut statics = Subtree::new();
    statics.insert(
        "write".into(),
        NamespaceNode::CommonEmit("dotnet.console_error_write".into()),
    );
    statics.insert(
        "writeline".into(),
        NamespaceNode::CommonEmit("dotnet.console_error_writeline".into()),
    );
    NamespaceNode::Namespace(statics)
}

#[cfg(test)]
mod resolve_gap_tests {
    use vybe_runtime::namespaces::{NamespaceNode, registry_read};

    /// These assert what this file is responsible for — that the entries are
    /// REGISTERED — rather than resolving through them. Resolution moved to
    /// `vybe_compiler::primitives::namespaces`, and a platform must not depend on
    /// the compiler; that dependency direction is the whole point of the
    /// plugin seam.
    fn registered_leaf(path: &[&str]) -> Option<NamespaceNode> {
        let guard = registry_read();
        let mut node = guard.tree.get(path[0])?.clone();
        for seg in &path[1..] {
            node = match node {
                NamespaceNode::Namespace(children) => children.get(*seg)?.clone(),
                NamespaceNode::Type {
                    statics, methods, ..
                } => statics.get(*seg).or_else(|| methods.get(*seg)).cloned()?,
                _ => return None,
            };
        }
        Some(node)
    }

    #[test]
    fn delegate_combine_is_registered() {
        super::register_namespace_tree();
        match registered_leaf(&["dotnet", "system", "delegate", "combine"]) {
            Some(NamespaceNode::CommonEmit(name)) => assert_eq!(name, "delegates.combine"),
            other => panic!("expected CommonEmit(delegates.combine), got {other:?}"),
        }
    }

    #[test]
    fn guid_parse_is_registered() {
        super::register_namespace_tree();
        assert!(
            registered_leaf(&["dotnet", "system", "guid", "parse"]).is_some(),
            "guid.parse not registered"
        );
    }
}

#[cfg(test)]
mod ctor_parity_tests {
    /// Every descriptor class with a backing constructor must resolve to the
    /// SAME target through the tree as through the old descriptor surface.
    #[test]
    fn tree_ctor_matches_descriptor_surface() {
        super::register_namespace_tree();
        let scope = super::dotnet_scope();
        let mut gaps = Vec::new();
        for export in crate::emitter::class_exports::dotnet_class_exports() {
            let Some(want) = export
                .class
                .constructor
                .as_ref()
                .and_then(|c| c.backing.clone())
            else {
                continue;
            };
            let got = vybe_runtime::namespaces::lookup_type_ctor_target(&scope, &export.class.name);
            if got.as_ref() != Some(&want) {
                gaps.push(format!(
                    "{}: want {:?} got {:?}",
                    export.class.name, want, got
                ));
            }
        }
        assert!(gaps.is_empty(), "{} gaps:\n{}", gaps.len(), gaps.join("\n"));
    }
}

#[cfg(test)]
mod member_parity_tests {
    /// Every descriptor INSTANCE method must resolve to the same target
    /// through the tree as through the descriptor surface. Any gap here is a
    /// hole the compiler falls through when the platform hook is gone.
    #[test]
    fn tree_instance_methods_match_descriptor_surface() {
        super::register_namespace_tree();
        let scope = super::dotnet_scope();
        let mut gaps = Vec::new();
        for export in crate::emitter::class_exports::dotnet_class_exports() {
            for m in &export.class.methods {
                if m.is_static {
                    continue;
                }
                let want = crate::emitter::surface().lookup_instance_method(
                    &export.class.name,
                    &m.name,
                    m.arity,
                );
                let got = vybe_runtime::namespaces::lookup_type_instance_target(
                    &scope,
                    &export.class.name,
                    &m.name,
                    m.arity,
                );
                if want.is_some() && got != want {
                    gaps.push(format!(
                        "{}.{}/{}: want {:?} got {:?}",
                        export.class.name, m.name, m.arity, want, got
                    ));
                }
            }
        }
        assert!(gaps.is_empty(), "{} gaps:\n{}", gaps.len(), gaps.join("\n"));
    }
}

#[cfg(test)]
mod property_parity_tests {
    /// Every descriptor INSTANCE property must resolve to the same target
    /// through the tree as through the descriptor surface.
    #[test]
    fn tree_instance_properties_match_descriptor_surface() {
        super::register_namespace_tree();
        let scope = super::dotnet_scope();
        let mut gaps = Vec::new();
        for export in crate::emitter::class_exports::dotnet_class_exports() {
            // Inherited too — `Button.Text` is declared on `Control`, and a
            // flat check would pass while every real read stayed broken.
            for p in super::inherited_properties(&export.class.name) {
                for want_setter in [false, true] {
                    let want = if want_setter {
                        crate::emitter::surface()
                            .lookup_instance_property_setter(&export.class.name, &p.name)
                    } else {
                        crate::emitter::surface()
                            .lookup_instance_property(&export.class.name, &p.name)
                    };
                    let got = if want_setter {
                        vybe_runtime::namespaces::lookup_type_property_setter_target(
                            &scope,
                            &export.class.name,
                            &p.name,
                        )
                    } else {
                        vybe_runtime::namespaces::lookup_type_property_target(
                            &scope,
                            &export.class.name,
                            &p.name,
                        )
                    };
                    if want.is_some() && got != want {
                        gaps.push(format!(
                            "{}.{}{}: want {:?} got {:?}",
                            export.class.name,
                            p.name,
                            if want_setter { " (set)" } else { "" },
                            want,
                            got
                        ));
                    }
                }
            }
        }
        assert!(gaps.is_empty(), "{} gaps:\n{}", gaps.len(), gaps.join("\n"));
    }
}
