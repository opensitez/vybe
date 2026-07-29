//! `dotnet.*` namespace-tree registration (namespaceplan.md, dotnet phase).
//!
//! The dotnet platform contributes DATA — its component-model class
//! descriptors — to the shared namespace tree in `vybe_bytecode::namespaces`.
//! Resolution LOGIC lives only in the common resolver: VB, C#, and every
//! other language resolve `dotnet.system.console.writeline` through the
//! same tree walk, instead of a platform-owned dotted-name cascade
//! (`resolver.rs`, which this registration supersedes and which dissolves
//! once VB/C# routing is fully migrated).

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_bytecode::component_model::{ConstructorTarget, MethodBody};
use vybe_bytecode::namespaces::{self, NamespaceNode, Subtree};

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
            for m in &class.methods {
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
            for p in inherited_properties(&class.name) {
                let node = namespaces::property(
                    p.getter.as_ref().map(|t| accessor_node(t, &p.name)),
                    p.setter.as_ref().map(|t| accessor_node(t, &p.name)),
                );
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

            let ty = NamespaceNode::Type {
                ctor: None,
                ctor_call,
                statics,
                methods,
                member_returns,
            };

            // "dotnet.System" + "Math" → dotnet.system.math
            let mut segments: Vec<String> =
                interface.split('.').map(|s| s.to_lowercase()).collect();
            segments.push(class.name.to_lowercase());

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

/// One accessor leaf. The generic `vybe:gui` property host functions take the
/// property NAME as an argument, so those bind it; a dedicated per-property
/// host function (`Environment.NewLine` → `node:os.EOL`) does not.
fn accessor_node(target: &vybe_bytecode::component_model::HostTarget, prop: &str) -> NamespaceNode {
    if target.name == vybe_compiler::compiler::gui::HOST_FN_GET_PROPERTY
        || target.name == vybe_compiler::compiler::gui::HOST_FN_SET_PROPERTY
    {
        namespaces::host_fn_keyed(&target.module, &target.name, prop)
    } else {
        namespaces::host_fn(&target.module, &target.name)
    }
}

/// A class's own properties followed by every inherited one, nearest first, so
/// an `or_insert` fold gives override-shadows-base.
fn inherited_properties(class_name: &str) -> Vec<vybe_bytecode::component_model::PropertyDef> {
    let descriptor = crate::emitter::surface().component_descriptor();
    let mut out = Vec::new();
    let mut current = descriptor
        .classes
        .iter()
        .find(|c| c.name.eq_ignore_ascii_case(class_name));
    let mut seen: Vec<String> = Vec::new();
    while let Some(class) = current {
        if seen.iter().any(|n| n.eq_ignore_ascii_case(&class.name)) {
            break; // cyclic parent chain — refuse to spin
        }
        seen.push(class.name.clone());
        out.extend(class.properties.iter().cloned());
        current = class.parent.as_deref().and_then(|parent| {
            descriptor
                .classes
                .iter()
                .find(|c| c.name.eq_ignore_ascii_case(parent))
        });
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
        _ => vec![],
    };
    entries
        .iter()
        .map(|(n, node)| ((*n).to_string(), node.clone()))
        .collect()
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
    use vybe_bytecode::namespaces::{registry_read, NamespaceNode};

    /// These assert what this file is responsible for — that the entries are
    /// REGISTERED — rather than resolving through them. Resolution moved to
    /// `vybe_compiler::compiler::namespaces`, and a platform must not depend on
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
            let got =
                vybe_bytecode::namespaces::lookup_type_ctor_target(&scope, &export.class.name);
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
                let got = vybe_bytecode::namespaces::lookup_type_instance_target(
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
                        vybe_bytecode::namespaces::lookup_type_property_setter_target(
                            &scope,
                            &export.class.name,
                            &p.name,
                        )
                    } else {
                        vybe_bytecode::namespaces::lookup_type_property_target(
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
