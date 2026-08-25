//! ⛔ KEYS KEEP THE DECLARED SPELLING. They used to be lowercased here, which
//! only worked because every tree lookup lowercased its query too. Lookups now
//! match EXACT first and fold only on a miss, so a case-sensitive language
//! resolves by the real name and a case-insensitive one still resolves by the
//! fold. See `documentation/casesensitivityplan.md`.
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
//! - the element it is built as is the `create` static leaf — GCL's Delphi
//!   surface is literally
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

/// The generic-construction spec for a GCL class. `control_fn` is the element
/// the class IS; the ancestry is what `is`/`isInstance` answer from.
///
/// **`params`/`fields`/`field_gui` are deliberately EMPTY, and that is the whole
/// difference between the two families of frontend.** A VCL control is
/// constructed BARE and configured afterwards by property assignment —
/// `TEdit.Create(Self)` takes an owner, not a `PasswordChar` — so there is no
/// constructor argument that names a property, and nothing to declare here.
/// WinForms is the same shape (`platforms/dotnet`, `control_ctor_spec`). Only a
/// DECLARATIVE frontend, where config arrives as constructor arguments, has
/// fields to describe: Flutter's `Text('hi', style: …)`.
///
/// Declaring the property list here instead aligned the ctor's ONE real
/// argument onto `properties[0]`, so `TMemo.Create(Self)` stamped
/// `PasswordChar = <the owner>` plus a `Count` of `null` that then raced the
/// real `Lines.Count` property read. Inert only because nothing on this path
/// consumes `__ops` — `emit_tree_ctor_construction` emits it today.
///
/// Nothing is lost: the MEMBERS come from the ancestry walk below, not from
/// this list. And `describes_construction` (`primitives/expressions.rs`) is
/// `!params.is_empty() || !fields.is_empty() || control_fn.is_some()`, so a
/// widget still routes to generic construction on `control_fn` alone — a class
/// with NO element would fall through to the backing-call path, which is why
/// this is safe here and would not be everywhere.
///
/// ⚠ A class with NO element keeps the old declaration, and that is forced, not
/// chosen. `describes_construction` reads non-emptiness as "construct this
/// generically", so emptying the lists for `TObject`/`TComponent`/`TControl`/
/// `TWinControl` — which have no `control_fn` to carry the disjunct — drops them
/// to the backing-call path, and they have no backing call. Measured:
/// `TComponent.Create(nil)` became `undefined is not callable`. Their
/// declaration is inert anyway (no `control_fn`, so no `__ops`; the stamped
/// fields are shadowed by the declared property members), so the hazard this
/// removes is entirely on the widget side. The real fix is a shared one — a
/// spec needs to say "generic construction, no fields" without saying it with
/// an empty list — and it is not plib's to make.
fn ctor_spec(class: &super::gcl::GclClass) -> CtorSpec {
    let Some(control_fn) = class.widget_host_fn else {
        let fields: Vec<String> = class.properties.iter().map(|p| p.to_string()).collect();
        return CtorSpec {
            params: fields.clone(),
            field_gui: fields
                .iter()
                .map(|f| FieldGui::NestOrProp(f.clone()))
                .collect(),
            fields,
            ancestry: ancestry(class),
            control_fn: None,
            // Not a control, so there is no chrome to be born with.
            inner_html: None,
            // A VCL control IS its element from the moment it is constructed —
            // there is no configuration/element split to bridge.
            nest_coerce: None,
            value_equality: false,
        };
    };
    CtorSpec {
        params: Vec::new(),
        fields: Vec::new(),
        field_gui: Vec::new(),
        ancestry: ancestry(class),
        control_fn: Some(control_fn.to_string()),
        // plib's composites (`TPageControl`, `TSplitter`) can declare their
        // default children here the same way dotnet's do; none does yet.
        inner_html: None,
        nest_coerce: None,
        value_equality: false,
    }
}

/// A VCL property spelling → the shared GUI ROLE it fills.
///
/// The roles are the shared canonical property names — the vocabulary every
/// language already spoke. This is the whole of plib's job here: `Caption` and
/// `Text` are the same role, `ClientWidth` is `width`. The Pascal word stops
/// at this function; nothing downstream knows VCL exists.
fn gui_property_role(owner: &str, prop: &str) -> &'static str {
    let spelling = prop.to_ascii_lowercase();
    // Spellings VCL reuses for unrelated things. The DECLARING class settles
    // them — `Position` on a form is where the window opens, on a track bar it
    // is the value — which is why this takes the owner rather than matching on
    // the word alone. A word-only map answered `Position` the same for both and
    // there was no spelling that could have told them apart.
    match (owner, spelling.as_str()) {
        // Window placement (`poScreenCenter`). No DOM counterpart — the page
        // does not choose where its window opens — so it stays an attribute.
        ("TForm", "position") => return "",
        (_, "position") => return "value",
        // `Lines.Count`, and only a memo's. A list control declaring `Count`
        // would want its item count, which is a different question.
        ("TMemo", "count") => return "linecount",
        _ => {}
    }
    match spelling.as_str() {
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
        "tag" => "tag",
        "scrollbars" => "overflow",
        // A spin edit's range and step. HTML already has all three on
        // `<input type=number>`, spelled `min` / `max` / `step`, and a track
        // bar declares them under those very names — so this is a SPELLING
        // difference and nothing more. Left unmapped they wrote
        // `minvalue="0"`, an attribute no element has ever read.
        "minvalue" => "min",
        "maxvalue" => "max",
        "increment" => "step",
        // VCL's `Color` is the control's BACKGROUND — `Font.Color` is the text.
        // WinForms spells the same two `BackColor`/`ForeColor`, which is what
        // the roles are named after.
        "color" => "backcolor",
        // `Font` IS a style property, reached like `Left` or `Color`. Every
        // `.dfm` in the corpus declares `Font.Name = 'Segoe UI'`, and without
        // this row the whole axis was dropped before it could reach the
        // cascade — which is why `vybe_widgets`' CSS inheritance measured
        // neutral despite `font_family`/`font_size`/`font_weight`/`font_style`
        // all being in its inherited set.
        "font" => "font",
        // VCL's `Alignment` IS CSS `text-align` — how the caption sits inside
        // the control's own box, which is what every framework means by it
        // (WinForms spells it `TextAlign`). The `ta*` constants are declared in
        // the Pascal profile as the CSS keywords, the same treatment `TAlign`
        // and `TScrollStyle` already get, so nothing downstream translates a
        // VCL enum.
        "alignment" | "textalign" => "textalign",
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

    vec![
        ("tlist", array_ctor()),
        ("tobjectlist", array_ctor()),
        ("tdictionary", map_type()),
    ]
}

/// Register the GCL class table under the `plib` root. Idempotent; first
/// call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut classes = Subtree::new();
        for class in super::gcl::gcl_classes() {
            // NO `create` STATIC LEAF. `TButton.Create` is CONSTRUCTION, and
            // construction is declared once — by `ctor_spec`, whose `control_fn`
            // carries the element and which the shared resolver drives through
            // `emit_control_element`. A leaf here declared the same fact a second
            // time as a raw HOST TARGET, and `widget_host_fn` is an HTML TAG
            // (`fieldset`, `select`, `input:checkbox`), so it named a host
            // function that does not exist. It only ever looked harmless because
            // the working path shadowed it.
            let statics = Subtree::new();
            // INSTANCE PROPERTIES. The class is an ADAPTER: `TEdit` is a
            // textbox, `Caption` is its text — the control and the concept are
            // already the same thing, so all a VCL class contributes is the
            // NAME. Declaring each property as a two-target member is what
            // lets the shared resolver answer `lbl.Caption := x` without the
            // compiler knowing Pascal exists; the generic accessors carry the
            // property name as a bound argument.

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
            for (owner, prop) in
                declared().flat_map(|c| c.properties.iter().map(move |p| (c.name, p)))
            {
                // Registered under the CANONICAL name, not the VCL spelling:
                // `Caption` IS `Text`, so Pascal and C# reach one property on
                // one control. The Pascal spelling stays the map KEY (that is
                // what source says); only the bound target is canonical.
                // The role this VCL property fills. Unmapped names keep their
                // own spelling, which lands on an attribute of that name.
                let role = match gui_property_role(owner, prop) {
                    "" => prop.to_ascii_lowercase(),
                    r => r.to_string(),
                };
                let canonical = role.as_str();
                if let Some(value_type) = gui::property_value_type(canonical) {
                    member_returns.insert(prop.to_string(), value_type.to_string());
                }
                members.insert(
                    prop.to_string(),
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
                    .entry(method.name.to_string())
                    .or_insert_with(|| match method.target {
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
                let key = member.to_string();
                members.insert(
                    key.clone(),
                    namespaces::property(
                        Some(NamespaceNode::CommonEmit("pascal.self".to_string())),
                        None,
                    ),
                );
                member_returns.insert(key, returns.to_string());
            }
            // `Items[i]` — the option list by index. See `INDEXED_ITEM_CLASSES`
            // for why the member is named `Item` and why both directions are
            // declared together. It reads back as a string, which is what makes
            // `Txt := Src.Items[i]` type as text rather than as nothing.
            if super::gcl::INDEXED_ITEM_CLASSES
                .iter()
                .any(|indexed| indexed.eq_ignore_ascii_case(class.name))
            {
                members.insert(
                    "item".to_string(),
                    namespaces::property(
                        Some(NamespaceNode::CommonEmit(gui::ITEM_TEXT_EMIT.to_string())),
                        Some(NamespaceNode::CommonEmit(
                            gui::SET_ITEM_TEXT_EMIT.to_string(),
                        )),
                    ),
                );
                // Pascal declares the indexed property as `Items`/`Item` with a
                // capital; the lowercase key was written for a folding tree.
                member_returns.insert("Item".to_string(), "string".to_string());
            }
            classes.insert(
                class.name.to_string(),
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
        // this tree, so the common resolver answers it. Building it imperatively
        // instead leaves `Application.Initialize` reading `undefined` off a
        // global nothing installed.
        //
        // `Title` IS `document.title` — the VCL word for the same thing HTML
        // already names, so it needs no host function of its own.
        let mut application = Subtree::new();
        // `Application.Run` EMITS NOTHING. A document is not told to run — it
        // runs because it HAS content, which is the condition the launch gate
        // already reads. `Run` was the last thing keeping a `should_run` flag
        // alive on a host that is being deleted (guiplan.md, "There is no
        // `runApplication`, for anybody").
        application.insert(
            "run".to_string(),
            NamespaceNode::CommonEmit(gui::APP_RUN_EMIT.to_string()),
        );
        application.insert(
            "terminate".to_string(),
            NamespaceNode::CommonEmit(gui::APP_EXIT_EMIT.to_string()),
        );
        // `Title` IS `document.title`, which is the `windowtitle` ROLE — the
        // same one a Form's caption fills. Declared as the role, so the VCL
        // word maps onto a concept HTML already names.
        application.insert(
            "title".to_string(),
            namespaces::property(
                Some(NamespaceNode::CommonEmit(format!(
                    "{}windowtitle",
                    gui::PROP_GET_EMIT
                ))),
                Some(NamespaceNode::CommonEmit(format!(
                    "{}windowtitle",
                    gui::PROP_SET_EMIT
                ))),
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
