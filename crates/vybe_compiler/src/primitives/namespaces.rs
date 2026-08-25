//! Namespace RESOLUTION — the compiler's query surface over the registry.
//!
//! The registry itself (types, storage, `register_namespace_tree`, the node
//! constructors) lives in `vybe_runtime::namespaces`, the crate every writer
//! already depends on. Resolution is compiler behaviour and lives here; the
//! POLICY around it — ambient roots, alias handling — is in
//! `crate::primitives::resolver`.
//!
//! Verified when this split was made: platforms and languages only ever WRITE.
//! Neither calls `resolve_path`.

use vybe_runtime::Value;
// Re-exported so `crate::primitives::namespaces::{CtorSpec, FieldGui, …}` keeps
// resolving for the compiler sites that consume the resolved shapes.
pub use vybe_runtime::namespaces::{CtorSpec, FieldGui, NamespaceNode, Path, Subtree};
use vybe_runtime::namespaces::{mount_host_exports, registry_read};

use super::*;

/// Maximum transitive `Alias` dereferences before resolution gives up —
/// guards registration mistakes that create alias cycles.
const MAX_ALIAS_DEPTH: usize = 8;

pub fn resolve_path(
    segments: &[&str],
    fold: vybe_runtime::namespaces::Fold,
) -> Option<ResolutionTarget> {
    if segments.is_empty() {
        return None;
    }
    mount_host_exports(
        crate::primitives::instructions::host::CapabilityContext::get()
            .functions
            .entries()
            .map(|(m, n)| (m.to_string(), n.to_string())),
    );
    let guard = registry_read();
    resolve_segments(&guard.tree, segments, 0, fold)
}

/// Walk an EXPLICIT forest instead of the process-global registry.
///
/// The `user.<unit>.*` root of namespaceplan.md lives on the `Compiler` rather
/// than in `vybe_runtime::namespaces` (per-unit mount vs. process-global tree),
/// but it must resolve by the same rules as `dotnet.*` or `ecma.*`. Sharing
/// `resolve_segments` is what makes the user root one more root in ONE
/// resolver rather than a second resolver wearing a tree's clothes.
pub(crate) fn resolve_in(
    forest: &Subtree,
    segments: &[&str],
    fold: vybe_runtime::namespaces::Fold,
) -> Option<ResolutionTarget> {
    if segments.is_empty() {
        return None;
    }
    resolve_segments(forest, segments, 0, fold)
}

pub(crate) fn normalize_source_path(name: &str) -> String {
    name.trim_start_matches(['\\', '.'])
        .split(['\\', '.'])
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>()
        .join(".")
}

/// Runtime twin of [`normalize_source_path`] — stack: `[string] → [string]`.
///
/// `normalize_source_path` canonicalizes a name the compiler can SEE. When the
/// name is only known at runtime (`$fn = "\\App\\helper"; $fn();`) the same
/// collapse has to happen in emitted code. Without it a call site has no way to
/// match the needle against the canonical `defined_functions` corpus, and the
/// only remaining option is the inverse: expand every KNOWN name back into
/// every spelling it could have been written with, and compare against all of
/// them. That is quadratic in emitted bytecode and re-introduces separator
/// syntax the frontend already normalized away.
///
/// Every spelling this collapses —
/// `A.B.f`, `A\B\f`, `A\\B\\f`, and each of those rooted with `\` or `\\` —
/// lands on one of exactly TWO keys: the canonical dotted name, or that name
/// with a leading `.` (see [`rooted_lookup_key`]) when the source rooted it.
/// Rooted-ness is preserved rather than stripped because it is meaningful:
/// `\strlen` explicitly opts out of the current namespace.
pub fn emit_normalize_source_path(chunk: &mut Chunk, line: u32) {
    // `\\` → `\` — a doubled separator is what an escaped literal leaves behind.
    chunk.emit_string_const("\\\\", line);
    chunk.emit_string_const("\\", line);
    crate::primitives::strings::emit_replace(chunk, line);

    // `\` → `.` — one dotted spelling regardless of the source separator.
    chunk.emit_string_const("\\", line);
    chunk.emit_string_const(".", line);
    crate::primitives::strings::emit_replace(chunk, line);
}

/// The key [`emit_normalize_source_path`] yields for a ROOTED spelling of an
/// already-canonical dotted name. The unrooted spelling yields the name itself.
pub fn rooted_lookup_key(canonical: &str) -> String {
    format!(".{canonical}")
}

impl Compiler {
    /// Resolve a source-level function name to the global identity produced by
    /// `NamespaceDecl` lowering.
    ///
    /// Frontends normalize their spelling (`\A\B\f`, `A.B.f`, imported alias,
    /// bare `f`) into an identifier. The compiler then applies one common
    /// policy:
    ///
    /// 1. exact dotted identity wins,
    /// 2. root-qualified bare names opt out of the enclosing namespace,
    /// 3. everything else is the `user.<unit>.*` tree — the same three tiers a
    ///    TYPE resolves through (exact, enclosing namespace innermost-first,
    ///    each imported prefix), asked of a declared FUNCTION.
    pub(crate) fn resolve_namespaced_function_identity(&self, name: &str) -> Option<String> {
        if !self.profile.uses_common_resolver {
            return None;
        }

        let rooted = name.starts_with(['\\', '.']);
        let normalized = normalize_source_path(name);
        if normalized.is_empty() {
            return None;
        }

        let exact = self.canon(&normalized);
        if self.defined_functions.contains(&exact) {
            return Some(exact);
        }

        // Root-qualified bare names (`\strlen`, `.println`) explicitly opt out
        // of the current namespace and should proceed as the global/builtin
        // spelling even when they are not user-defined functions.
        if rooted && !normalized.contains('.') {
            return Some(exact);
        }

        // TWO flat tiers stood here and both are gone: a `current_namespace`
        // probe of `defined_functions`, then `resolve_source_namespace_value`,
        // which scans that same set for the exact spelling, then each enclosing
        // namespace, then each import. Every answer either was a strict subset
        // of the tree walk below or was the identical answer reached by
        // scanning a flat set instead of walking scope.
        //
        // Subsets, precisely: the `current_namespace` probe consulted ONE
        // namespace and never popped a segment, where `source_namespace_contexts`
        // supplies both the enclosing `Namespace` and the namespace of the class
        // being compiled — not the same thing, since a method body compiles with
        // `current_namespace` already unwound — and the tree walks them
        // innermost-first. `resolve_source_namespace_value` used those same
        // contexts and the same `source_namespace_imports`, in the same order,
        // over `defined_functions ∪ defined_globals`, then filtered the result
        // back down to `defined_functions` — so its `defined_globals` arm could
        // never produce an answer here at all.
        //
        // They cannot disagree with the tree about membership either: the whole
        // program has ONE `defined_functions.insert`, inside
        // `declare_function_identity`, which registers the tree leaf in the same
        // breath. A name is in both or in neither.
        //
        // The `user.<unit>.*` root — the same three tiers (exact spelling,
        // enclosing namespace innermost-first, each imported prefix) that a
        // TYPE resolves through, now asked of a declared FUNCTION.
        //
        // What stood here was the shape types were rescued from: a
        // `ends_with(".{name}")` scan of `defined_functions` accepting a
        // unique match, then a bare last-segment lookup. Both are guesses over
        // a flat set — they answer by coincidence of spelling rather than by
        // scope, so a namespace that was never imported still resolved, and
        // two namespaces declaring the same member name made the answer depend
        // on how many happened to exist.
        if let Some(qualified) = self.resolve_user_namespace_function(&normalized) {
            return Some(qualified);
        }

        None
    }
}

fn resolve_segments(
    forest: &Subtree,
    segments: &[&str],
    alias_depth: usize,
    fold: vybe_runtime::namespaces::Fold,
) -> Option<ResolutionTarget> {
    if alias_depth > MAX_ALIAS_DEPTH {
        return None; // alias cycle or absurd chain — refuse, don't spin
    }
    let mut current: &Subtree = forest;
    let mut walked: Vec<&str> = Vec::with_capacity(segments.len());

    for (i, seg) in segments.iter().enumerate() {
        // ⛔ THE SAME RULE THE TREE USES — exact first, fold only on a miss.
        // A second lookup rule here is how the resolver and the tree come to
        // disagree about what resolves: this walker sees the SOURCE's spelling
        // (`Encoding`, `getBytes`) and the tree now holds the REGISTRAR's, and
        // only one function is allowed to decide whether those are the same
        // name.
        let node = vybe_runtime::namespaces::fold_get(current, seg, fold)?;
        walked.push(seg);
        let is_last = i + 1 == segments.len();

        if is_last {
            return terminal(forest, node, &walked.join("."), alias_depth, fold);
        }

        // Descend.
        match node {
            NamespaceNode::Namespace(children) => current = children,
            NamespaceNode::Type { statics, .. } => current = statics,
            // A user declaration is a namespace over what is declared under
            // it, the same way a `Type` is a namespace over its statics.
            NamespaceNode::UserGlobal { children, .. } => current = children,
            NamespaceNode::Alias(target) => {
                // Re-root the walk at the alias target with the remaining
                // segments appended.
                let mut re_rooted: Vec<&str> = target.split('.').collect();
                re_rooted.extend_from_slice(&segments[i + 1..]);
                return resolve_segments(forest, &re_rooted, alias_depth + 1, fold);
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
    /// A declaration of the unit being compiled, resolved through the
    /// `user.<unit>.*` root. `identity` is the canonical global name every
    /// downstream table keys on.
    UserGlobal {
        identity: Path,
        kind: vybe_runtime::namespaces::UserGlobalKind,
    },
}

fn terminal(
    forest: &Subtree,
    node: &NamespaceNode,
    path: &str,
    alias_depth: usize,
    fold: vybe_runtime::namespaces::Fold,
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
            .and_then(|n| terminal(forest, n, path, alias_depth, fold)),
        NamespaceNode::Overloads(entries) => entries
            .last()
            .and_then(|(_, n)| terminal(forest, n, path, alias_depth, fold)),
        NamespaceNode::Const(v) => Some(ResolutionTarget::Const(v.clone())),
        // The unit's own declaration. Deliberately NOT a `Ctor`: construction
        // of a user class goes through the class machinery keyed on
        // `identity`, and answering `Ctor` here would also make a declared
        // FUNCTION look constructible.
        NamespaceNode::UserGlobal { identity, kind, .. } => {
            Some(ResolutionTarget::UserGlobal {
                identity: identity.clone(),
                kind: *kind,
            })
        }
        NamespaceNode::Alias(target) => {
            let segs: Vec<&str> = target.split('.').collect();
            resolve_segments(forest, &segs, alias_depth + 1, fold)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;
    use vybe_runtime::namespaces::*;

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
        match resolve_path(&["ecma", "json", "stringify"], vybe_runtime::namespaces::FOLD_ASCII) {
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
        match resolve_path(&["python", "json", "dumps"], vybe_runtime::namespaces::FOLD_ASCII) {
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
        match resolve_path(&["ecma", "json"], vybe_runtime::namespaces::FOLD_ASCII) {
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
        match resolve_path(&["python", "js_like", "json", "parse"], vybe_runtime::namespaces::FOLD_ASCII) {
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
        assert!(resolve_path(&["a", "x"], vybe_runtime::namespaces::FOLD_ASCII).is_none());
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
        assert!(resolve_path(&["ecma", "json", "parse"], vybe_runtime::namespaces::FOLD_ASCII).is_some());
        assert!(resolve_path(&["ecma", "math", "max"], vybe_runtime::namespaces::FOLD_ASCII).is_some());
    }

    #[test]
    fn type_statics_and_ctor() {
        let _g = LOCK.lock().unwrap();
        clear_registry_for_tests();
        let mut statics = Subtree::new();
        statics.insert("writeline".into(), host_fn("web:console", "log"));
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
            resolve_path(&["dotnet", "system", "console"], vybe_runtime::namespaces::FOLD_ASCII),
            Some(ResolutionTarget::Ctor { .. })
        ));
        assert!(matches!(
            resolve_path(&["dotnet", "system", "console", "writeline"], vybe_runtime::namespaces::FOLD_ASCII),
            Some(ResolutionTarget::HostCall { .. })
        ));
    }

    #[test]
    fn unknown_paths_fall_through() {
        let _g = LOCK.lock().unwrap();
        seed();
        assert!(resolve_path(&["nope"], vybe_runtime::namespaces::FOLD_ASCII).is_none());
        assert!(resolve_path(&["ecma", "nope"], vybe_runtime::namespaces::FOLD_ASCII).is_none());
        // Leaf with trailing segments is not a namespace path.
        assert!(resolve_path(&["ecma", "json", "stringify", "extra"], vybe_runtime::namespaces::FOLD_ASCII).is_none());
    }
}
