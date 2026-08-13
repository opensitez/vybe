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
            let class_is_control = element_backed_control(&class.name);
            // The WHOLE inherited surface, not just this class's own methods —
            // the tree has no parent link, so anything left off here is
            // unreachable. See `inherited_methods`.
            for m in &inherited_methods(&class.name) {
                // A control METHOD is a shared VERB, resolved through
                // `primitives/gui.rs` like its properties — not a `vybe:gui`
                // host call. Gated on the class actually being a control for
                // the same reason the accessors are: `Graphics` and the value
                // types declare methods too, and `Update`/`Refresh` are
                // ordinary names that must not be captured off a non-control.
                let gui_verb = if class_is_control && !m.is_static {
                    gui_control_verb(&m.name)
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
                    let entries = method_overloads.entry(m.name.to_lowercase()).or_default();
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
                let entries = bucket.entry(m.name.to_lowercase()).or_default();
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
            for p in inherited_properties(&class.name) {
                let node = namespaces::property(
                    p.getter
                        .as_ref()
                        .map(|t| accessor_node(t, &p.name, is_control)),
                    p.setter
                        .as_ref()
                        .map(|t| accessor_node(t, &p.name, is_control)),
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
                        property_returns.insert(p.name.to_lowercase(), value_type.to_string());
                    }
                }
                methods
                    .entry(p.name.to_lowercase())
                    .or_insert_with(|| node.clone());
                statics.entry(p.name.to_lowercase()).or_insert(node);
            }
            for (name, node) in shared_emit_accessors(&class.name) {
                methods.insert(name, node);
            }
            if interface.eq_ignore_ascii_case("dotnet.System")
                && class.name.eq_ignore_ascii_case("Console")
            {
                statics.insert("out".into(), console_stdout_writer_node());
                statics.insert("error".into(), console_stderr_writer_node());
            }
            if interface.eq_ignore_ascii_case("dotnet.System")
                && class.name.eq_ignore_ascii_case("Object")
            {
                statics.insert(
                    "equals".into(),
                    NamespaceNode::CommonEmit("dotnet.object_equals".into()),
                );
                statics.insert(
                    "referenceequals".into(),
                    NamespaceNode::CommonEmit("dotnet.object_reference_equals".into()),
                );
            }
            if interface.eq_ignore_ascii_case("dotnet.System")
                && class.name.eq_ignore_ascii_case("DateTime")
            {
                statics.insert(
                    "minvalue".into(),
                    NamespaceNode::CommonEmit("dotnet.datetime_min_value".into()),
                );
                statics.insert(
                    "maxvalue".into(),
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
                    member_returns.insert(m.name.to_lowercase(), rt);
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
                    .entry(name.to_lowercase())
                    .or_insert_with(|| (*ty).to_string());
            }
            // A member that IS the receiver overrides both — the descriptor's
            // own signature describes the collection wrapper WinForms models,
            // and the element does not have one.
            for (name, ty) in self_member_returns(&class.name) {
                member_returns.insert((*name).to_string(), (*ty).to_string());
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
            let ctor = html_element_for_control(&class.name)
                .map(|element| control_ctor_spec(&class.name, element));

            let ty = NamespaceNode::Type {
                ctor,
                ctor_call,
                statics,
                methods,
                member_returns,
            };

            // "dotnet.System" + "Math" → dotnet.system.math
            let mut segments: Vec<String> =
                interface.split('.').map(|s| s.to_lowercase()).collect();
            segments.push(class.name.to_lowercase());

            if interface.eq_ignore_ascii_case("dotnet.System") {
                namespaces::register_namespace_tree(&class.name.to_lowercase(), ty.clone());
            }

            let mut node = ty;
            while segments.len() > 1 {
                let key = segments.pop().expect("non-empty");
                let mut children = Subtree::new();
                children.insert(key, node);
                node = NamespaceNode::Namespace(children);
            }
            namespaces::register_namespace_tree(&segments.pop().expect("root"), node);
        }
    });
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
/// close to the identity: `vybe:gui`'s property vocabulary was modelled on
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
/// `primitives/gui.rs` instead of each calling its own `vybe:gui` host fn.
///
/// `Refresh`/`Invalidate`/`Update` map to real verbs that lower to NOTHING —
/// a document repaints itself, so there is nothing for an author to ask for.
/// That is deliberate and is explained at `emit_gui_control_method`.
fn gui_control_verb(method: &str) -> Option<&'static str> {
    Some(match method.to_ascii_lowercase().as_str() {
        "show" => "show",
        "hide" => "hide",
        "focus" => "focus",
        "refresh" => "refresh",
        "invalidate" => "invalidate",
        "update" => "update",
        // ⛔ `BringToFront`/`SendToBack` are NOT converted yet, on purpose.
        // Z-order is document order, so they are `appendChild` /
        // `insertBefore(firstChild)` — but `web:dom` exports neither
        // `parentNode`, `insertBefore` nor `firstChild` today (it has
        // `appendChild`/`removeChild` and the attribute/text surface, nothing
        // for reading a parent or a first child). All three are standard DOM
        // and belong in `web:*`; adding them is a `platforms/web` decision.
        //
        // Until then these keep their existing `vybe:gui` target, which WORKS.
        // Mapping them now would trade a working call for an
        // `Unresolved import` — and lowering them to a no-op would be a silent
        // shim, which is worse than either.
        _ => return None,
    })
}

/// One accessor leaf.
///
/// The two generic `vybe:gui` property host functions take the property NAME as
/// an argument, so they are the CONTROL property path — and they are what this
/// platform is being converted off. They now bind to the shared role emits
/// (`gui.prop_get.<role>` / `gui.prop_set.<role>`) that `primitives/gui.rs`
/// lowers to `web:dom` / `web:html` / `web:cssom`, which is the same target
/// plib already reaches. Under the hood both still drive `vybe_widgets`; only
/// the route changed, from a bespoke host function to a DOM operation.
///
/// A dedicated per-property host function (`Environment.NewLine` →
/// `node:os.EOL`) is NOT a control property and is left exactly as it was.
///
/// `is_control` is what keeps the value types out. `Point`/`Size`/`Font` bind
/// the same two generic host functions, so the target name alone cannot tell
/// them apart from a `Button` — only the class can, and it does it by whether
/// it descends from `Control`.
fn accessor_node(
    target: &vybe_runtime::component_model::HostTarget,
    prop: &str,
    is_control: bool,
) -> NamespaceNode {
    let setting = target.name == vybe_compiler::primitives::gui::HOST_FN_SET_PROPERTY;
    if is_control
        && (setting || target.name == vybe_compiler::primitives::gui::HOST_FN_GET_PROPERTY)
    {
        let role = match gui_property_role(prop) {
            "" => prop.to_ascii_lowercase(),
            r => r.to_string(),
        };
        let prefix = if setting {
            vybe_compiler::primitives::gui::PROP_SET_EMIT
        } else {
            vybe_compiler::primitives::gui::PROP_GET_EMIT
        };
        NamespaceNode::CommonEmit(format!("{prefix}{role}"))
    } else if target.name == vybe_compiler::primitives::gui::HOST_FN_GET_PROPERTY
        || target.name == vybe_compiler::primitives::gui::HOST_FN_SET_PROPERTY
    {
        // Not a control: keep the keyed host-function accessor exactly as it
        // was. `vybe:gui` still answers for the value types and the drawing
        // surface, which is why this arm survives the control conversion.
        namespaces::host_fn_keyed(&target.module, &target.name, prop)
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
        // maps to it, which is `vybe_widgets`' half of the mapping, not this
        // registrar's.
        "statusstrip" => "footer",
        "combobox" => "select",
        "listbox" => "ul",
        "treeview" => "ul",
        // HTML has these outright, and they carry real semantics a `<div>`
        // cannot: a range input is keyboard-operable and `<progress>` is
        // announced as a progress indicator.
        "progressbar" => "progress",
        "trackbar" => "input:range",
        "numericupdown" => "input:number",

        // ── No HTML counterpart: a DECLARED custom element ─────────────────
        // Declaring `vybe-*` explicitly is better than letting it fall through
        // to `ControlElement::custom`, because the choice becomes visible here
        // instead of being implied by absence — and it keeps these names in
        // step with plib, which spells the same controls the same way
        // (`TImage`, `TSplitter`, `TPageControl`, `TTabSheet`, `TTimer`).
        "picturebox" => "vybe-picturebox",
        // The scrollbars and the navigator. `vybe_widgets` has had all three
        // kinds and their default sizes the whole time; what was missing was
        // the DECLARATION, without which they had no `CtorSpec` and every
        // property write on them was silently dropped.
        // ⚠ A ContextMenuStrip is NOT `<menu>`, and the difference is docking.
        // The `menu` TAG is born `Dock::Top`, which is right for a menu bar and
        // wrong for a popup: `cms1` took the full width under the other strips
        // and threw away the `Location`/`Size` the designer gave it, pushing
        // the first real control off the top of the form. A context menu is
        // attached to a control and shown on demand — it is not a bar.
        "contextmenustrip" => "vybe-contextmenustrip",
        // Declared `vybe-*` custom elements. `control_kind` strips `vybe-` and
        // looks the remainder up against the widget list, so the TAG carries
        // the kind and these two land on real widgets that already exist
        // (`checkedlistbox`; `datagrid` folds onto `datagridview`).
        "checkedlistbox" => "vybe-checkedlistbox",
        "datagrid" => "vybe-datagrid",
        // ⚠ These have no widget kind YET, so they render as a label until
        // `vybe_widgets` grows one — the designed degradation, visible in a
        // capture instead of the control vanishing. The declaration still buys
        // construction, identity, geometry, text, events and data binding.
        "propertygrid" => "vybe-propertygrid",
        "splitter" => "vybe-splitter",
        "domainupdown" => "vybe-domainupdown",
        // A UserControl is a plain composite container, and `<section>` is a
        // real element that already establishes a containing block.
        "usercontrol" => "section",
        "hscrollbar" => "vybe-hscrollbar",
        "vscrollbar" => "vybe-vscrollbar",
        "bindingnavigator" => "vybe-bindingnavigator",
        // ⚠ `vybe-splitter`, not `vybe-splitcontainer`, made this a LABEL.
        // The tag carries the kind: `control_kind` strips `vybe-` and looks the
        // remainder up against the widget list, which spells it
        // `splitcontainer`. A tag naming no known control degrades to a 120x20
        // label — so a mapping that renames the control silently deletes it.
        // plib's `TSplitter` is a different control (a bare drag bar); this one
        // is WinForms' two-panel container.
        "splitcontainer" => "vybe-splitcontainer",
        "tabcontrol" => "vybe-tabcontrol",
        "tabpage" => "vybe-tabpage",
        "timer" => "vybe-timer",

        // ── Deliberately NOT converted yet ─────────────────────────────────
        // `Panel`/`FlowLayoutPanel`/`TableLayoutPanel` are `<div>` and plib
        // already maps `TPanel` that way — but `<div>` containers currently
        // lay out wrong in the shared `vybe_widgets` engine (children carry
        // scaled CSS and never get a laid-out rect, while body-parented
        // siblings do). That is a shared layout defect, not a mapping
        // question, so importing the mapping now would import the bug.
        // Convert these once it is fixed.
        //
        // `ListView`/`DataGridView`/`WebBrowser` have no counterpart at all;
        // forcing them onto `<table>` would claim semantics they do not have.
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
        value_equality: false,
    }
}

/// Does this class DESCEND FROM `Control` — i.e. is it a control at all?
///
/// The element model applies to controls and nothing else. `Point`, `Size`,
/// `Font`, `Pen`, `Brush` and `Graphics` are value types and a drawing surface:
/// they are bound to `vybe:gui` constructors for historical reasons but they
/// are not elements, have no DOM counterpart, and must keep their host-function
/// accessors until canvas/CSSOM answers for them separately.
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

/// Does this class map to an element? The descriptor builder asks, so a class
/// with no `vybe:gui` factory still gets a constructor — see
/// `winforms/component_classes.rs`.
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
            ("item", rw("dotnet.sb_index_get", "dotnet.sb_index_set")),
            ("length", rw("dotnet.sb_length", "dotnet.sb_set_length")),
            (
                "capacity",
                rw("dotnet.sb_capacity", "dotnet.sb_set_capacity"),
            ),
            ("maxcapacity", ro("dotnet.sb_max_capacity")),
        ],
        "stopwatch" => vec![
            ("elapsedmilliseconds", ro("dotnet.stopwatch_elapsed_ms")),
            ("elapsedticks", ro("dotnet.stopwatch_elapsed_ticks")),
            ("elapsed", ro("dotnet.stopwatch_elapsed")),
            ("isrunning", ro("dotnet.stopwatch_is_running")),
        ],
        "task" => vec![
            ("result", ro("dotnet.task_result")),
            ("iscompleted", ro("dotnet.task_is_completed")),
            ("iscanceled", ro("dotnet.task_is_canceled")),
        ],
        "list" | "arraylist" => vec![("capacity", ro("dotnet.list_capacity"))],
        // The two members a cursor cannot store — both are derived from
        // whatever `DataSource` currently points at, so a field would go stale
        // the moment the source changed. `Position`, `DataMember`, `Filter` and
        // `Sort` are NOT here: those are real fields the constructor writes.
        "bindingsource" => vec![
            ("count", ro("dotnet.bindingsource_count")),
            ("current", ro("dotnet.bindingsource_current")),
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
            ("items", ro("dotnet.self")),
            ("dropdownitems", ro("dotnet.self")),
            (
                "add",
                namespaces::overloads(vec![(
                    2,
                    emit(vybe_compiler::primitives::gui::APPEND_CHILD_EMIT),
                )]),
            ),
        ],
        _ => vec![],
    };
    entries
        .iter()
        .map(|(n, node)| ((*n).to_string(), node.clone()))
        .collect()
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
            ("items", "ToolStripMenuItem"),
            ("dropdownitems", "ToolStripMenuItem"),
        ],
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
