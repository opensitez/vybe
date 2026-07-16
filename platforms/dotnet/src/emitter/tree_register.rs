//! `dotnet.*` namespace-tree registration (namespaceplan.md, dotnet phase).
//!
//! The dotnet platform contributes DATA — its component-model class
//! descriptors — to the shared namespace tree in `vybe_emitter::namespaces`.
//! Resolution LOGIC lives only in the common resolver: VB, C#, and every
//! other language resolve `dotnet.system.console.writeline` through the
//! same tree walk, instead of a platform-owned dotted-name cascade
//! (`resolver.rs`, which this registration supersedes and which dissolves
//! once VB/C# routing is fully migrated).

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_bytecode::component_model::MethodBody;
use vybe_emitter::namespaces::{self, NamespaceNode, Subtree};

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
            for m in &class.methods {
                if !m.is_static {
                    // Instance methods dispatch receiver-based (TypeRegistry
                    // vtables), never through the namespace tree.
                    continue;
                }
                let node = match &m.body {
                    MethodBody::Common(emit) => NamespaceNode::CommonEmit(emit.clone()),
                    MethodBody::HostCall(t) => namespaces::host_fn(&t.module, &t.name),
                    // Chunk-backed methods are per-compilation artifacts,
                    // not process-global surface.
                    MethodBody::UserChunk(_) => continue,
                };
                statics.insert(m.name.to_lowercase(), node);
            }
            // Host-backed property getters — the legacy static-property
            // lookup surface (`Stopwatch.Frequency`-style reads through a
            // class path) as tree leaves. Methods win on name collision.
            for p in &class.properties {
                if let Some(getter) = &p.getter {
                    statics
                        .entry(p.name.to_lowercase())
                        .or_insert_with(|| namespaces::host_fn(&getter.module, &getter.name));
                }
            }
            let ty = NamespaceNode::Type {
                ctor: None,
                statics,
                methods: BTreeMap::new(),
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

#[cfg(test)]
mod resolve_gap_tests {
    use vybe_emitter::namespaces::{resolve_path, ResolutionTarget};

    #[test]
    fn delegate_combine_resolves_via_tree() {
        super::register_namespace_tree();
        let r = resolve_path(&["dotnet", "system", "delegate", "combine"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "delegates.combine"),
            other => panic!("expected CommonEmit(delegates.combine), got {other:?}"),
        }
    }

    #[test]
    fn guid_parse_resolves_via_tree() {
        super::register_namespace_tree();
        let r = resolve_path(&["dotnet", "system", "guid", "parse"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.guid_parse"),
            other => panic!("expected CommonEmit(dotnet.guid_parse), got {other:?}"),
        }
    }
}
