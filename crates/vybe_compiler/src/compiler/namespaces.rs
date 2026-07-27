//! Namespace RESOLUTION — the compiler's query surface over the registry.
//!
//! The registry itself (types, storage, `register_namespace_tree`, the node
//! constructors) lives in `vybe_bytecode::namespaces`, the crate every writer
//! already depends on. Resolution is compiler behaviour and lives here; the
//! POLICY around it — ambient roots, alias handling — is in
//! `crate::compiler::resolver`.
//!
//! Verified when this split was made: platforms and languages only ever WRITE.
//! Neither calls `resolve_path`.

use vybe_bytecode::Value;
// Re-exported so `crate::compiler::namespaces::{CtorSpec, FieldGui, …}` keeps
// resolving for the compiler sites that consume the resolved shapes.
pub use vybe_bytecode::namespaces::{CtorSpec, FieldGui, NamespaceNode, Path, Subtree};
use vybe_bytecode::namespaces::{mount_host_exports, registry_read};

/// Maximum transitive `Alias` dereferences before resolution gives up —
/// guards registration mistakes that create alias cycles.
const MAX_ALIAS_DEPTH: usize = 8;

pub fn resolve_path(segments: &[&str]) -> Option<ResolutionTarget> {
    if segments.is_empty() {
        return None;
    }
    mount_host_exports(
        crate::compiler::instructions::host::CapabilityContext::get()
            .functions
            .entries()
            .map(|(m, n)| (m.to_string(), n.to_string())),
    );
    let guard = registry_read();
    resolve_segments(&guard.tree, segments, 0)
}

fn resolve_segments(
    forest: &Subtree,
    segments: &[&str],
    alias_depth: usize,
) -> Option<ResolutionTarget> {
    if alias_depth > MAX_ALIAS_DEPTH {
        return None; // alias cycle or absurd chain — refuse, don't spin
    }
    let mut current: &Subtree = forest;
    let mut walked: Vec<&str> = Vec::with_capacity(segments.len());

    for (i, seg) in segments.iter().enumerate() {
        let node = current.get(*seg)?;
        walked.push(seg);
        let is_last = i + 1 == segments.len();

        if is_last {
            return terminal(forest, node, &walked.join("."), alias_depth);
        }

        // Descend.
        match node {
            NamespaceNode::Namespace(children) => current = children,
            NamespaceNode::Type { statics, .. } => current = statics,
            NamespaceNode::Alias(target) => {
                // Re-root the walk at the alias target with the remaining
                // segments appended.
                let mut re_rooted: Vec<&str> = target.split('.').collect();
                re_rooted.extend_from_slice(&segments[i + 1..]);
                return resolve_segments(forest, &re_rooted, alias_depth + 1);
            }
            // A function/const leaf with segments still remaining is not a
            // namespace path.
            // A function/const/overload-set leaf with segments still
            // remaining is not a namespace path.
            NamespaceNode::Fn { .. }
            | NamespaceNode::CommonEmit(_)
            | NamespaceNode::Const(_)
            | NamespaceNode::Overloads(_)
            | NamespaceNode::Property { .. } => {
                return None;
            }
        }
    }
    None
}


#[derive(Debug, Clone)]
pub enum ResolutionTarget {
    /// Direct host call: `CALL_IMPORT module/func`.
    HostCall {
        module: String,
        func: String,
        /// Arity when the registrar knew it. Every consumer currently ignores
        /// this (`HostCall { module, func, .. }`); it exists so descriptor-
        /// registered methods can be dispatched by arity without the compiler
        /// asking a platform crate.
        arity: Option<u8>,
    },
    /// `common:<cat>.<op>` dispatch.
    CommonEmit(String),
    /// A constructable type at `path`, with the spec that drives generic
    /// construction (named-arg reorder + field capture + `is` ancestry).
    Ctor { path: Path, spec: Option<CtorSpec> },
    /// Compile-time constant, inlined at use-site.
    Const(Value),
    /// The path names a namespace (or type used as a namespace) — callers
    /// materialize an ECMA-262 §16.2 module-namespace object if a runtime
    /// value is required.
    NamespaceObject(Path),
}

fn terminal(
    forest: &Subtree,
    node: &NamespaceNode,
    path: &str,
    alias_depth: usize,
) -> Option<ResolutionTarget> {
    match node {
        NamespaceNode::Namespace(_) => Some(ResolutionTarget::NamespaceObject(path.to_string())),
        NamespaceNode::Type { ctor, .. } => Some(ResolutionTarget::Ctor {
            path: path.to_string(),
            spec: ctor.clone(),
        }),
        NamespaceNode::Fn {
            module,
            func,
            arity,
            ..
        } => Some(ResolutionTarget::HostCall {
            module: module.clone(),
            func: func.clone(),
            arity: *arity,
        }),
        NamespaceNode::CommonEmit(name) => Some(ResolutionTarget::CommonEmit(name.clone())),
        // A path walk carries no argument count, so it cannot discriminate
        // overloads; resolve to the LAST-declared target, which is what a
        // name-keyed map left behind before overloads were tree data.
        // Call sites that know their argc go through
        // `lookup_type_instance_target`, which selects properly.
        // A property read through a path walk resolves to its GETTER — the
        // walk is a read. Writes go through the setter lookup, which knows
        // the direction.
        NamespaceNode::Property { get, .. } => get
            .as_deref()
            .and_then(|n| terminal(forest, n, path, alias_depth)),
        NamespaceNode::Overloads(entries) => entries
            .last()
            .and_then(|(_, n)| terminal(forest, n, path, alias_depth)),
        NamespaceNode::Const(v) => Some(ResolutionTarget::Const(v.clone())),
        NamespaceNode::Alias(target) => {
            let segs: Vec<&str> = target.split('.').collect();
            resolve_segments(forest, &segs, alias_depth + 1)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_bytecode::namespaces::*;
    use std::sync::Mutex;

    // The registry is process-global; serialize tests that touch it.
    static LOCK: Mutex<()> = Mutex::new(());

    fn seed() {
        clear_registry_for_tests();
        register_namespace_tree(
            "ecma",
            namespace(vec![(
                "json",
                namespace(vec![
                    ("stringify", host_fn("ecma:json", "stringify")),
                    ("parse", host_fn("ecma:json", "parse")),
                ]),
            )]),
        );
        register_namespace_tree(
            "python",
            namespace(vec![(
                "json",
                namespace(vec![
                    ("dumps", NamespaceNode::Alias("ecma.json.stringify".into())),
                    ("loads", NamespaceNode::Alias("ecma.json.parse".into())),
                ]),
            )]),
        );
    }

    #[test]
    fn resolves_direct_host_leaf() {
        let _g = LOCK.lock().unwrap();
        seed();
        match resolve_path(&["ecma", "json", "stringify"]) {
            Some(ResolutionTarget::HostCall { module, func, .. }) => {
                assert_eq!(module, "ecma:json");
                assert_eq!(func, "stringify");
            }
            other => panic!("expected HostCall, got {:?}", other),
        }
    }

    #[test]
    fn resolves_alias_to_canonical_leaf() {
        let _g = LOCK.lock().unwrap();
        seed();
        match resolve_path(&["python", "json", "dumps"]) {
            Some(ResolutionTarget::HostCall { module, func, .. }) => {
                assert_eq!(module, "ecma:json");
                assert_eq!(func, "stringify");
            }
            other => panic!("expected HostCall via alias, got {:?}", other),
        }
    }

    #[test]
    fn interior_path_is_namespace_object() {
        let _g = LOCK.lock().unwrap();
        seed();
        match resolve_path(&["ecma", "json"]) {
            Some(ResolutionTarget::NamespaceObject(p)) => assert_eq!(p, "ecma.json"),
            other => panic!("expected NamespaceObject, got {:?}", other),
        }
    }

    #[test]
    fn alias_namespace_descends_into_target() {
        let _g = LOCK.lock().unwrap();
        seed();
        // Alias at an interior position: python.js_like → ecma, then walk on.
        register_namespace_tree("python", {
            namespace(vec![("js_like", NamespaceNode::Alias("ecma".into()))])
        });
        match resolve_path(&["python", "js_like", "json", "parse"]) {
            Some(ResolutionTarget::HostCall { func, .. }) => assert_eq!(func, "parse"),
            other => panic!("expected HostCall through interior alias, got {:?}", other),
        }
    }

    #[test]
    fn alias_cycle_is_refused() {
        let _g = LOCK.lock().unwrap();
        clear_registry_for_tests();
        register_namespace_tree(
            "a",
            namespace(vec![("x", NamespaceNode::Alias("b.y".into()))]),
        );
        register_namespace_tree(
            "b",
            namespace(vec![("y", NamespaceNode::Alias("a.x".into()))]),
        );
        assert!(resolve_path(&["a", "x"]).is_none());
    }

    #[test]
    fn merge_preserves_sibling_registrations() {
        let _g = LOCK.lock().unwrap();
        seed();
        // A second registration under "ecma" must not clobber ecma.json.
        register_namespace_tree(
            "ecma",
            namespace(vec![(
                "math",
                namespace(vec![("max", host_fn("ecma:math", "max"))]),
            )]),
        );
        assert!(resolve_path(&["ecma", "json", "parse"]).is_some());
        assert!(resolve_path(&["ecma", "math", "max"]).is_some());
    }

    #[test]
    fn type_statics_and_ctor() {
        let _g = LOCK.lock().unwrap();
        clear_registry_for_tests();
        let mut statics = Subtree::new();
        statics.insert("writeline".into(), host_fn("wasi:logging/logging", "log"));
        register_namespace_tree(
            "dotnet",
            namespace(vec![(
                "system",
                namespace(vec![(
                    "console",
                    NamespaceNode::Type {
                        ctor: None,
                        ctor_call: None,
                        statics,
                        methods: Subtree::new(),
                        member_returns: Default::default(),
                    },
                )]),
            )]),
        );
        assert!(matches!(
            resolve_path(&["dotnet", "system", "console"]),
            Some(ResolutionTarget::Ctor { .. })
        ));
        assert!(matches!(
            resolve_path(&["dotnet", "system", "console", "writeline"]),
            Some(ResolutionTarget::HostCall { .. })
        ));
    }

    #[test]
    fn unknown_paths_fall_through() {
        let _g = LOCK.lock().unwrap();
        seed();
        assert!(resolve_path(&["nope"]).is_none());
        assert!(resolve_path(&["ecma", "nope"]).is_none());
        // Leaf with trailing segments is not a namespace path.
        assert!(resolve_path(&["ecma", "json", "stringify", "extra"]).is_none());
    }
}

