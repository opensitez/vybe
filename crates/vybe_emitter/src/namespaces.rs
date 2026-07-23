//! Namespace tree — the single global name-resolution MODEL (namespaceplan.md).
//!
//! One rooted forest shared by ALL languages. Roots are platform/language
//! packages (`ecma.*`, `wasi.*`, `node.*`, `web.*`, `dotnet.*`, `php.*`,
//! `libc.*`, `plib.*`, `user.<unit>.*`, …). Any language may reference any
//! root by fully-qualified path; a language's profile *mounts* decide only
//! what is ambient, not what is reachable.
//!
//! Leaves are **component-model typed** (`FuncSig` from
//! `vybe_bytecode::component`) — resolution ends at a declared, typed
//! interface member, not a bare string pair. `namespaces.rs` is *meaning*;
//! `imports.rs` stays *mechanism* (chunk-level `add_import`/CALL_IMPORT).
//!
//! ## Case canon
//!
//! Tree keys are stored **lowercase-canonical**. Case-sensitive languages
//! canon at the walker/compiler boundary; the resolver itself NEVER
//! lowercases (the `matchAll` bug class — see namespaceplan.md gotchas).
//!
//! ## Aliases (source-name ≠ canonical-name)
//!
//! Name divergence is pervasive, not rare: Python `json.dumps` →
//! `ecma.json.stringify`, PHP `json_encode` → the same leaf. `Alias` nodes
//! point at a leaf living at a different path; resolution dereferences them
//! transitively (cycle-guarded) to a terminal typed leaf.
//!
//! ## Phase 0 scope
//!
//! Model + process-global registry only — NO consumer behavior change. The
//! compiler-core resolver (`vybe_compiler::compiler::resolver`) is the sole
//! query surface; languages migrate one at a time (JS → Python → PHP →
//! dotnet → rest), each keeping its suite green.

use std::collections::BTreeMap;
use std::sync::{OnceLock, RwLock};

use vybe_bytecode::Value;
use vybe_bytecode::component::FuncSig;

/// A dot-separated canonical path into the tree (`"ecma.json.stringify"`).
pub type Path = String;

/// Children of a namespace (or the statics of a type).
pub type Subtree = BTreeMap<String, NamespaceNode>;

/// One node in the global namespace forest.
#[derive(Debug, Clone)]
pub enum NamespaceNode {
    /// An interior namespace: `ecma`, `ecma.json`, `dotnet.system`.
    Namespace(Subtree),
    /// A class/type — CM-typed constructor + static members + instance
    /// method signatures. Instance methods are listed for *metadata* only:
    /// member dispatch is receiver-based (TypeRegistry vtables) and NEVER
    /// resolves through the namespace tree; only `ctor` and `statics` are
    /// reachable by path walk.
    Type {
        ctor: Option<FuncSig>,
        statics: Subtree,
        methods: BTreeMap<String, FuncSig>,
    },
    /// A host-backed typed function leaf: resolves to a `CALL_IMPORT` of
    /// `module`/`func` (e.g. `ecma:json` / `stringify`).
    Fn {
        module: String,
        func: String,
        sig: FuncSig,
    },
    /// A `common:<cat>.<op>` dispatch leaf — emitted through the shared
    /// common-emit dispatcher rather than a direct host call.
    CommonEmit(String),
    /// A compile-time constant value (`Math.PI`-class leaves).
    Const(Value),
    /// This name → a leaf living at a different canonical path. The
    /// source≠canonical reconciliation point (`python.json.dumps` →
    /// `Alias("ecma.json.stringify")`).
    Alias(Path),
}

/// What a successful path walk resolves to.
#[derive(Debug, Clone)]
pub enum ResolutionTarget {
    /// Direct host call: `CALL_IMPORT module/func`.
    HostCall {
        module: String,
        func: String,
        sig: FuncSig,
    },
    /// `common:<cat>.<op>` dispatch.
    CommonEmit(String),
    /// A constructable type at `path`.
    Ctor { path: Path, sig: Option<FuncSig> },
    /// Compile-time constant, inlined at use-site.
    Const(Value),
    /// The path names a namespace (or type used as a namespace) — callers
    /// materialize an ECMA-262 §16.2 module-namespace object if a runtime
    /// value is required.
    NamespaceObject(Path),
}

/// Maximum transitive `Alias` dereferences before resolution gives up —
/// guards registration mistakes that create alias cycles.
const MAX_ALIAS_DEPTH: usize = 8;

static REGISTRY: OnceLock<RwLock<Subtree>> = OnceLock::new();

fn registry() -> &'static RwLock<Subtree> {
    REGISTRY.get_or_init(|| RwLock::new(BTreeMap::new()))
}

/// Register (or merge) a tree under `root`. Process-global data, intended
/// to be immutable once startup registration finishes — per-VM *mounts*
/// are a separate concern and never mutate the tree.
///
/// Merging: when `root` already exists and both the existing node and the
/// new one are `Namespace`s, children merge recursively (platforms
/// register their surface in independent pieces); any other collision
/// replaces the previous node (later registration wins, mirroring host
/// binary-loader registration).
///
/// Keys must already be lowercase-canonical — see the module docs.
pub fn register_namespace_tree(root: &str, tree: NamespaceNode) {
    debug_assert_eq!(
        root,
        root.to_lowercase(),
        "namespace tree keys are stored lowercase-canonical; canon at the boundary"
    );
    let mut guard = registry().write().unwrap();
    merge_into(&mut guard, root.to_string(), tree);
}

fn merge_into(map: &mut Subtree, key: String, node: NamespaceNode) {
    match (map.get_mut(&key), node) {
        (Some(NamespaceNode::Namespace(existing)), NamespaceNode::Namespace(new_children)) => {
            for (k, v) in new_children {
                merge_into(existing, k, v);
            }
        }
        (Some(NamespaceNode::Type { statics, .. }), NamespaceNode::Namespace(new_children)) => {
            for (k, v) in new_children {
                merge_into(statics, k, v);
            }
        }
        (
            Some(existing @ NamespaceNode::Namespace(_)),
            NamespaceNode::Type {
                ctor,
                mut statics,
                methods,
            },
        ) => {
            if let NamespaceNode::Namespace(existing_children) = existing {
                let mut merged = existing_children.clone();
                for (k, v) in statics {
                    merge_into(&mut merged, k, v);
                }
                statics = merged;
            }
            *existing = NamespaceNode::Type {
                ctor,
                statics,
                methods,
            };
        }
        (_, node) => {
            map.insert(key, node);
        }
    }
}

/// Resolve a canonical, already-canonned path through the forest.
///
/// Returns `None` when the path names nothing — the caller falls through
/// to ordinary member/receiver dispatch (resolver step 4).
///
/// The host component-model surface mounts lazily on first query, from the
/// SAME `FunctionRegistry` the emitter validates host calls against — one
/// source of truth for what the host exports, no VM or runtime involved.
pub fn resolve_path(segments: &[&str]) -> Option<ResolutionTarget> {
    if segments.is_empty() {
        return None;
    }
    mount_host_exports(
        crate::instructions::host::CapabilityContext::get()
            .functions
            .entries()
            .map(|(m, n)| (m.to_string(), n.to_string())),
    );
    let guard = registry().read().unwrap();
    resolve_segments(&guard, segments, 0)
}

/// True when `root` has been registered — cheap membership probe so
/// callers can distinguish "unknown root" from "known root, bad path".
pub fn has_root(root: &str) -> bool {
    registry().read().unwrap().contains_key(root)
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
            NamespaceNode::Fn { .. } | NamespaceNode::CommonEmit(_) | NamespaceNode::Const(_) => {
                return None;
            }
        }
    }
    None
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
            sig: ctor.clone(),
        }),
        NamespaceNode::Fn { module, func, sig } => Some(ResolutionTarget::HostCall {
            module: module.clone(),
            func: func.clone(),
            sig: sig.clone(),
        }),
        NamespaceNode::CommonEmit(name) => Some(ResolutionTarget::CommonEmit(name.clone())),
        NamespaceNode::Const(v) => Some(ResolutionTarget::Const(v.clone())),
        NamespaceNode::Alias(target) => {
            let segs: Vec<&str> = target.split('.').collect();
            resolve_segments(forest, &segs, alias_depth + 1)
        }
    }
}

/// Test-only: wipe the registry so unit tests are order-independent.
#[doc(hidden)]
pub fn clear_registry_for_tests() {
    registry().write().unwrap().clear();
    HOST_EXPORTS_MOUNTED.store(false, std::sync::atomic::Ordering::SeqCst);
}

static HOST_EXPORTS_MOUNTED: std::sync::atomic::AtomicBool =
    std::sync::atomic::AtomicBool::new(false);

/// Mount the host's registered export surface into the tree:
/// `("ecma:json", "stringify")` → `Fn` leaf at `ecma.json.stringify`,
/// `("wasi:cli/stdout", "write-via-stream")` → `wasi.cli.stdout.…`.
///
/// Called at startup wherever a VM exists (vybex runtime service, test
/// harnesses) with `vm.iter_host_function_exports()` pairs. First mount
/// wins — the host surface is identical across VMs in a process, so
/// remounting is skipped (`register_namespace_tree` stays available for
/// platform/alias registration, which is additive).
///
/// Keys are lowercased for storage (tree canon); the leaf payload keeps
/// the host's true casing (`isArray`) so emission never mangles the
/// name — the `matchAll` bug class.
pub fn mount_host_exports<I>(exports: I)
where
    I: IntoIterator<Item = (String, String)>,
{
    if HOST_EXPORTS_MOUNTED.swap(true, std::sync::atomic::Ordering::SeqCst) {
        return;
    }
    let mut guard = registry().write().unwrap();
    for (module, func) in exports {
        // "ecma:json" → [ecma, json]; "wasi:cli/stdout" → [wasi, cli, stdout]
        let mut segments: Vec<String> = Vec::new();
        let mut rest = module.as_str();
        if let Some((pkg, tail)) = rest.split_once(':') {
            segments.push(pkg.to_lowercase());
            rest = tail;
        }
        for part in rest.split('/') {
            if !part.is_empty() {
                segments.push(part.to_lowercase());
            }
        }
        if segments.is_empty() {
            continue;
        }
        segments.push(func.to_lowercase());

        // Build the nested single-child tree for this leaf, then merge.
        let mut node = NamespaceNode::Fn {
            module: module.clone(),
            func: func.clone(),
            sig: FuncSig {
                name: func.clone(),
                params: vec![],
                results: vec![],
            },
        };
        while segments.len() > 1 {
            let key = segments.pop().unwrap();
            let mut children = Subtree::new();
            children.insert(key, node);
            node = NamespaceNode::Namespace(children);
        }
        merge_into(&mut guard, segments.pop().unwrap(), node);
    }
}

// ── Construction helpers ────────────────────────────────────────────────

/// An untyped host-fn leaf (`Any` params/results) — the common case until
/// platforms register real CM signatures.
pub fn host_fn(module: &str, func: &str) -> NamespaceNode {
    NamespaceNode::Fn {
        module: module.to_string(),
        func: func.to_string(),
        sig: FuncSig {
            name: func.to_string(),
            params: vec![],
            results: vec![],
        },
    }
}

/// A namespace from a list of children.
pub fn namespace(children: Vec<(&str, NamespaceNode)>) -> NamespaceNode {
    NamespaceNode::Namespace(
        children
            .into_iter()
            .map(|(k, v)| {
                debug_assert_eq!(
                    k,
                    k.to_lowercase(),
                    "namespace tree keys are stored lowercase-canonical"
                );
                (k.to_string(), v)
            })
            .collect(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
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
                        statics,
                        methods: BTreeMap::new(),
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
