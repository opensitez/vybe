//! `plib.*` namespace-tree registration (namespaceplan.md: "Pascal GCL
//! surface").
//!
//! Mirrors the dotnet registrar: the plib platform contributes DATA — its
//! GCL class table (`gcl::gcl_classes()`, the same table the Pascal
//! lowering executes) — to the shared namespace tree in
//! `vybe_runtime::namespaces`. Resolution LOGIC lives only in the common
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
use vybe_runtime::namespaces::{self, CtorSpec, FieldGui, NamespaceNode, Subtree};

/// The GCL row for a class name, however it was spelled.
fn gcl_class(name: &str) -> Option<&'static super::gcl::GclClass> {
    super::gcl::gcl_classes()
        .iter()
        .find(|c| c.name.eq_ignore_ascii_case(name))
}

/// The declared parent of a GCL class name — the one link
/// [`namespaces::ancestry_of`] needs to walk this table.
fn gcl_parent(name: &str) -> Option<String> {
    gcl_class(name).and_then(|c| c.parent).map(str::to_string)
}

/// The `is`/`inherits` ancestry of a GCL class — self first, then its parent
/// chain, so `TButton is TControl` answers from the same `__types` stamp every
/// other adapter uses.
fn ancestry(class: &super::gcl::GclClass) -> Vec<String> {
    namespaces::ancestry_of(class.name, gcl_parent)
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

/// A VCL property spelling → the shared GUI ROLE it fills.
///
/// The roles are `vybe:gui`'s canonical property names — the vocabulary every
/// language already spoke. This is the whole of plib's job here: `Caption` and
/// `Text` are the same role, `ClientWidth` is `width`. The Pascal word stops
/// at this function; nothing downstream knows VCL exists.
fn gui_property_role(prop: &str) -> &'static str {
    match prop.to_ascii_lowercase().as_str() {
        "caption" | "text" => "text",
        "name" => "name",
        "checked" | "state" => "checked",
        "enabled" => "enabled",
        "visible" => "visible",
        "left" => "left",
        "top" => "top",
        "width" | "clientwidth" => "width",
        "height" | "clientheight" => "height",
        "itemindex" => "selectedindex",
        "items" => "items",
        "readonly" => "readonly",
        "maxlength" => "maxlength",
        "hint" => "tooltip",
        // `Lines.Count` — how many lines the control's text holds. Only
        // `TMemo` declares `Count`; the day a list control does, this needs to
        // become a per-class answer rather than a spelling one.
        "count" => "linecount",
        // VCL's `Align` IS WinForms' `Dock` — the control gives up its own
        // rect and takes an edge of the container. The constants it is
        // assigned are declared below as the role's own vocabulary.
        "align" => "dock",
        // A form's menu bar IS a child element of the form. `Menu` and
        // `MainMenu` are the same assignment `FMainMenu.Items.Add` makes, only
        // spelled from the container's side — one role, one emit.
        // `PopupMenu` deliberately stays an attribute: it is attached but not
        // displayed, so inserting it would render a stray menu bar.
        "menu" | "mainmenu" => "child",
        // Anything else keeps its own spelling and lands on an attribute —
        // a declared property with no shared role is still a property.
        _ => "",
    }
}

/// Delphi's `Generics.Collections` types — CONSTRUCTION only.
///
/// A `TList` IS the shared ordered-sequence store, the same built-in .NET's
/// `List<T>`, Python's list and a JS Array name, so `TList.Create` is
/// `collections.new` and nothing else. Its MEMBERS are not here: `Add`,
/// `Count`, `Delete` are spellings, and a spelling is declared by the language
/// that speaks it (`[value_methods]` / `[builtin_types]` in the Pascal
/// profile), never bound to a lexical type name in shared data.
///
/// Construction is the one place the type name is genuinely the subject —
/// `TList.Create` names it in source — and `lookup_type_ctor_target` already
/// strips the generic argument list, so `TList<Integer>.Create` resolves here
/// too without anything downstream knowing what a generic is.
fn collection_ctors() -> Vec<(&'static str, NamespaceNode)> {
    let array_ctor = || NamespaceNode::Type {
        ctor: None,
        ctor_call: Some(Box::new(NamespaceNode::CommonEmit(
            "collections.new".to_string(),
        ))),
        statics: Subtree::new(),
        // The two members that are PROPERTY READS rather than calls. A read
        // has no argument list for a spelling table to match on, so it is the
        // one member kind that needs the receiver's declared type — `Count`
        // fills the shared Len slot, `Items` the GetItem/SetItem pair.
        methods: BTreeMap::from([
            (
                "add".to_string(),
                NamespaceNode::CommonEmit("collections.push".to_string()),
            ),
            (
                "insert".to_string(),
                NamespaceNode::CommonEmit("collections.insert".to_string()),
            ),
            (
                "delete".to_string(),
                NamespaceNode::CommonEmit("collections.remove_at".to_string()),
            ),
            (
                "clear".to_string(),
                NamespaceNode::CommonEmit("collections.clear".to_string()),
            ),
            (
                "remove".to_string(),
                NamespaceNode::CommonEmit("collections.remove".to_string()),
            ),
            (
                "contains".to_string(),
                NamespaceNode::CommonEmit("collections.contains".to_string()),
            ),
            (
                "indexof".to_string(),
                NamespaceNode::CommonEmit("collections.index_of".to_string()),
            ),
            (
                "lastindexof".to_string(),
                NamespaceNode::CommonEmit("collections.last_index_of".to_string()),
            ),
            (
                "reverse".to_string(),
                NamespaceNode::CommonEmit("collections.reverse".to_string()),
            ),
            (
                "sort".to_string(),
                NamespaceNode::CommonEmit("collections.sort".to_string()),
            ),
            (
                "toarray".to_string(),
                NamespaceNode::CommonEmit("collections.clone".to_string()),
            ),
            // `Extract`/`ExtractAt` ARE `list.pop(i)` — a shared concept with
            // a route already.
            // Delphi-only members: no shared concept, or an argument a
            // `CommonEmit` name cannot carry. They decompose into the same
            // `collections.*` routes inside the Pascal emitter.
            (
                "extractat".to_string(),
                NamespaceNode::CommonEmit("pascal.list_extract_at".to_string()),
            ),
            (
                "extract".to_string(),
                NamespaceNode::CommonEmit("pascal.list_extract".to_string()),
            ),
            (
                "first".to_string(),
                NamespaceNode::CommonEmit("pascal.list_first".to_string()),
            ),
            (
                "last".to_string(),
                NamespaceNode::CommonEmit("pascal.list_last".to_string()),
            ),
            (
                "exchange".to_string(),
                NamespaceNode::CommonEmit("pascal.list_exchange".to_string()),
            ),
            (
                "move".to_string(),
                NamespaceNode::CommonEmit("pascal.list_move".to_string()),
            ),
            (
                "addrange".to_string(),
                NamespaceNode::CommonEmit("pascal.list_add_range".to_string()),
            ),
            (
                "trimexcess".to_string(),
                NamespaceNode::CommonEmit("pascal.list_noop".to_string()),
            ),
            (
                "count".to_string(),
                namespaces::property(
                    Some(NamespaceNode::CommonEmit("collections.length".to_string())),
                    None,
                ),
            ),
            // `Items[i]` is an INDEXED property: the read hands back the store
            // itself and the subscript that follows does the rest. `L.Items[i]`
            // is therefore literally `L[i]` — which it has to be, since the
            // list IS the array. Binding the read to `collections.get` instead
            // would arrive one argument short.
            (
                "items".to_string(),
                namespaces::property(
                    Some(NamespaceNode::CommonEmit("pascal.self".to_string())),
                    None,
                ),
            ),
        ]),
        // The declared return type of each member. Not decoration: Pascal
        // resolves OVERLOADS by static argument type, so a member whose
        // return type is unknown picks the wrong one — `__vs(L.Contains(x))`
        // took the `integer` overload and printed `0` for `False`. The
        // prelude class carried these on its signatures; they have to be
        // declared somewhere, and this is the type's own description of
        // itself.
        member_returns: BTreeMap::from([
            ("contains".to_string(), "Boolean".to_string()),
            ("indexof".to_string(), "Integer".to_string()),
            ("lastindexof".to_string(), "Integer".to_string()),
            ("count".to_string(), "Integer".to_string()),
            ("remove".to_string(), "Integer".to_string()),
        ]),
    };
    // `TDictionary` IS the shared Map — the same store PHP's array and
    // Python's dict land on, so a Delphi dictionary handed to either is a
    // thing they already understand. Not two parallel key/value arrays, which
    // is what the synthesized prelude built.
    let map_type = || NamespaceNode::Type {
        ctor: None,
        ctor_call: Some(Box::new(NamespaceNode::CommonEmit(
            "pascal.dict_new".to_string(),
        ))),
        statics: Subtree::new(),
        methods: BTreeMap::from([
            // Reads and writes go through the polymorphic `ARRAY_GET`/`SET`,
            // which already dispatch on `ObjectKind` and so are Map-aware.
            // The members that ENUMERATE do not: `common:dict.*` is the older
            // Ordinary+`__keys` shape and answers 0/`[]` on a Map, silently.
            // See the note in `runtime_adapter.rs`.
            (
                "add".to_string(),
                NamespaceNode::CommonEmit("dict.set_dynamic".to_string()),
            ),
            (
                "addorsetvalue".to_string(),
                NamespaceNode::CommonEmit("dict.set_dynamic".to_string()),
            ),
            (
                "containskey".to_string(),
                NamespaceNode::CommonEmit("pascal.dict_has".to_string()),
            ),
            (
                "containsvalue".to_string(),
                NamespaceNode::CommonEmit("pascal.dict_contains_value".to_string()),
            ),
            (
                "remove".to_string(),
                NamespaceNode::CommonEmit("pascal.dict_delete".to_string()),
            ),
            (
                "clear".to_string(),
                NamespaceNode::CommonEmit("pascal.dict_clear".to_string()),
            ),
            (
                "toarray".to_string(),
                NamespaceNode::CommonEmit("pascal.dict_items".to_string()),
            ),
            (
                "count".to_string(),
                namespaces::property(
                    Some(NamespaceNode::CommonEmit("pascal.dict_size".to_string())),
                    None,
                ),
            ),
            (
                "keys".to_string(),
                namespaces::property(
                    Some(NamespaceNode::CommonEmit("pascal.dict_keys".to_string())),
                    None,
                ),
            ),
            (
                "values".to_string(),
                namespaces::property(
                    Some(NamespaceNode::CommonEmit("pascal.dict_values".to_string())),
                    None,
                ),
            ),
            // Indexed property, same shape as the list's — see there.
            (
                "items".to_string(),
                namespaces::property(
                    Some(NamespaceNode::CommonEmit("pascal.self".to_string())),
                    None,
                ),
            ),
        ]),
        // See the note on the list's — overload resolution reads these.
        member_returns: BTreeMap::from([
            ("containskey".to_string(), "Boolean".to_string()),
            ("containsvalue".to_string(), "Boolean".to_string()),
            ("trygetvalue".to_string(), "Boolean".to_string()),
            ("count".to_string(), "Integer".to_string()),
        ]),
    };

    // A queue and a stack are the SAME array store; they differ only in which
    // end each Delphi verb reaches. `Peek` is the one member that has to know.
    let fifo_lifo = |peek: &str, take: &str| NamespaceNode::Type {
        ctor: None,
        ctor_call: Some(Box::new(NamespaceNode::CommonEmit(
            "collections.new".to_string(),
        ))),
        statics: Subtree::new(),
        methods: BTreeMap::from([
            (
                "enqueue".to_string(),
                NamespaceNode::CommonEmit("collections.push".to_string()),
            ),
            (
                "push".to_string(),
                NamespaceNode::CommonEmit("collections.push".to_string()),
            ),
            (
                "dequeue".to_string(),
                NamespaceNode::CommonEmit(take.to_string()),
            ),
            (
                "pop".to_string(),
                NamespaceNode::CommonEmit(take.to_string()),
            ),
            (
                "peek".to_string(),
                NamespaceNode::CommonEmit(peek.to_string()),
            ),
            (
                "clear".to_string(),
                NamespaceNode::CommonEmit("collections.clear".to_string()),
            ),
            (
                "toarray".to_string(),
                NamespaceNode::CommonEmit("collections.clone".to_string()),
            ),
            (
                "count".to_string(),
                namespaces::property(
                    Some(NamespaceNode::CommonEmit("collections.length".to_string())),
                    None,
                ),
            ),
        ]),
        // See the note on the list's — overload resolution reads these.
        member_returns: BTreeMap::from([("count".to_string(), "Integer".to_string())]),
    };

    vec![
        ("tlist", array_ctor()),
        ("tobjectlist", array_ctor()),
        ("tdictionary", map_type()),
        (
            "tqueue",
            fifo_lifo("pascal.list_first", "collections.shift"),
        ),
        ("tstack", fifo_lifo("pascal.list_last", "collections.pop")),
    ]
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
            // INSTANCE PROPERTIES. The class is an ADAPTER: `TEdit` is a
            // textbox, `Caption` is its text — the control and the concept are
            // already the same thing, so all a VCL class contributes is the
            // NAME. Declaring each property as a two-target member is what
            // lets the shared resolver answer `lbl.Caption := x` without the
            // compiler knowing Pascal exists; the generic accessors carry the
            // property name as a bound argument.
            //
            // `methods` was left empty by every registrar, which is exactly
            // why the compiler used to reach into platform crates for this.
            let mut members: Subtree = BTreeMap::new();
            // What a property READS BACK as. Declared from the role, so the one
            // answer serves every frontend registered this way, and the
            // ordinary expression machinery can work on a property's value:
            // `(Sender as TButton).Caption[1]` is a string subscript, and
            // undeclared it read `null`.
            let mut member_returns: std::collections::BTreeMap<String, String> =
                std::collections::BTreeMap::new();
            // A control has the properties of its whole chain — `TLabel.Caption`
            // is declared on `TControl`. Expanded HERE, at registration, so the
            // resolver answers with a flat lookup and no ancestry walk: the
            // adapter knows its own inheritance, nothing downstream should.
            //
            // Properties and methods walk the SAME chain — one `ancestry_of`,
            // read twice. They were two hand-rolled walks that had drifted:
            // methods take the nearest declaration, properties the furthest.
            // That difference is invisible here only because a re-declared
            // spelling (`TForm.Caption` over `TControl.Caption`) resolves to
            // the same role either way, so both orders build the same node.
            let chain = ancestry(class);
            let declared = || chain.iter().filter_map(|name| gcl_class(name));
            for prop in declared().flat_map(|c| c.properties) {
                // Registered under the CANONICAL name, not the VCL spelling:
                // `Caption` IS `Text`, so Pascal and C# reach one property on
                // one control. The Pascal spelling stays the map KEY (that is
                // what source says); only the bound target is canonical.
                // The role this VCL property fills. Unmapped names keep their
                // own spelling, which lands on an attribute of that name.
                let role = match gui_property_role(prop) {
                    "" => prop.to_ascii_lowercase(),
                    r => r.to_string(),
                };
                let canonical = role.as_str();
                if let Some(value_type) = gui::property_value_type(canonical) {
                    member_returns.insert(prop.to_lowercase(), value_type.to_string());
                }
                members.insert(
                    prop.to_lowercase(),
                    namespaces::property(
                        Some(NamespaceNode::CommonEmit(format!(
                            "{}{}",
                            gui::PROP_GET_EMIT,
                            canonical
                        ))),
                        Some(NamespaceNode::CommonEmit(format!(
                            "{}{}",
                            gui::PROP_SET_EMIT,
                            canonical
                        ))),
                    ),
                );
            }
            // METHODS, from the same chain as the properties. `GclClass.methods`
            // had no consumer at all — `Add`/`Show`/`Close` were declared and
            // never registered, so `FMainMenu.Items.Add(x)` called `undefined`.
            for method in declared().flat_map(|c| c.methods) {
                // Nearest declaration wins, so a subclass may override.
                members
                    .entry(method.name.to_lowercase())
                    .or_insert_with(|| match method.target {
                        super::gcl::GclMethodTarget::Host { module, fn_name } => {
                            namespaces::host_fn(module, fn_name)
                        }
                        super::gcl::GclMethodTarget::Common { emit } => {
                            NamespaceNode::CommonEmit(emit.to_string())
                        }
                    });
            }
            // Members that ARE the control — see `SELF_MEMBERS` for why the
            // declared return type is the load-bearing half of this.
            for (_, member, returns) in super::gcl::SELF_MEMBERS
                .iter()
                .filter(|(owner, _, _)| owner.eq_ignore_ascii_case(class.name))
            {
                let key = member.to_lowercase();
                members.insert(
                    key.clone(),
                    namespaces::property(
                        Some(NamespaceNode::CommonEmit("pascal.self".to_string())),
                        None,
                    ),
                );
                member_returns.insert(key, returns.to_string());
            }
            classes.insert(
                class.name.to_lowercase(),
                NamespaceNode::Type {
                    ctor: Some(ctor_spec(class)),
                    ctor_call: None,
                    statics,
                    methods: members,
                    member_returns,
                },
            );
        }
        // `Application` — the VCL global. DECLARED, like everything else in
        // this tree, so the common resolver answers it; it used to be built
        // imperatively by a `gcl/builder.rs` that had NO CALLER, so
        // `Application.Initialize` read `undefined` off a global that was
        // never installed. That whole file is deleted.
        //
        // `Title` IS `document.title` — the VCL word for the same thing HTML
        // already names, so it needs no host function of its own.
        let mut application = Subtree::new();
        application.insert(
            "run".to_string(),
            namespaces::host_fn(gui::GUI_MODULE, gui::HOST_FN_RUN_APPLICATION),
        );
        application.insert(
            "terminate".to_string(),
            namespaces::host_fn(gui::GUI_MODULE, gui::HOST_FN_APP_EXIT),
        );
        application.insert(
            "title".to_string(),
            namespaces::property(
                Some(namespaces::host_fn(gui::DOCUMENT_MODULE, "title")),
                Some(namespaces::host_fn(gui::DOCUMENT_MODULE, "setTitle")),
            ),
        );
        classes.insert(
            "application".to_string(),
            NamespaceNode::Namespace(application),
        );

        for (name, node) in collection_ctors() {
            classes.insert(name.to_string(), node);
        }

        namespaces::register_namespace_tree("plib", NamespaceNode::Namespace(classes));
    });
}

#[cfg(test)]
mod tests {
    /// The registration must ANSWER — a declared VCL property has to resolve
    /// to its role, or every lookup downstream silently falls through.
    #[test]
    fn declared_properties_resolve_to_roles() {
        super::register_namespace_tree();
        let scope = vec!["plib".to_string()];
        for (class, prop) in [
            ("tcontrol", "Caption"),
            ("twincontrol", "Caption"),
            ("tlabel", "Caption"),
            ("tedit", "Text"),
            ("tform", "Name"),
        ] {
            let found = vybe_runtime::namespaces::lookup_type_instance_member(&scope, class, prop);
            assert!(
                found.is_some(),
                "{class}.{prop} did not resolve — registration is not answering"
            );
        }
        let target = vybe_runtime::namespaces::lookup_type_property_setter_target(
            &scope, "tlabel", "Caption",
        );
        assert!(target.is_some(), "Caption has no setter target");
        println!("Caption setter target = {target:?}");
    }
}
