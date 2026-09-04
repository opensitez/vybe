//! The namespace tree — REGISTRY and RESOLUTION, in the compiler.
//!
//! The tree is a COMPILE-TIME name-resolution structure: a plugin's
//! `tree_register` writes its surface here, and the compiler walks it to turn a
//! dotted source path into a host call or a common emit. Nothing executes
//! against it. A component's RUN-TIME surface is its imports and exports plus
//! the canonical ABI (`vybe_runtime::component`, `canon_*`), which is why a
//! dotted `dotnet.system.threading.thread.sleep` has no meaning to the VM.
//!
//! Write side (plugins): `register_namespace_tree`, plus the node constructors
//! `host_fn` / `namespace` / `property` / `overloads`.
//!
//! Read side (compiler): `resolve_path` and the typed `lookup_type_*` walks.
//! The POLICY around them — ambient roots, alias handling, resolution order —
//! is `crate::primitives::resolver`, the one resolver of namespaceplan.md.

use std::collections::BTreeMap;
use std::sync::{OnceLock, RwLock};

use vybe_runtime::Value;

use super::*;

/// Maximum transitive `Alias` dereferences before resolution gives up —
/// guards registration mistakes that create alias cycles.
const MAX_ALIAS_DEPTH: usize = 8;

pub fn resolve_path(
    segments: &[&str],
    fold: crate::primitives::namespaces::Fold,
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
/// than in `crate::primitives::namespaces` (per-unit mount vs. process-global tree),
/// but it must resolve by the same rules as `dotnet.*` or `ecma.*`. Sharing
/// `resolve_segments` is what makes the user root one more root in ONE
/// resolver rather than a second resolver wearing a tree's clothes.
pub(crate) fn resolve_in(
    forest: &Subtree,
    segments: &[&str],
    fold: crate::primitives::namespaces::Fold,
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
    fold: crate::primitives::namespaces::Fold,
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
        let node = crate::primitives::namespaces::fold_get(current, seg, fold)?;
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
        /// The leaf's declared Component Model signature, when the registrar
        /// stated one. The RESOLVED TARGET carries it, so a named-argument
        /// call reorders against the leaf it just resolved to — never against
        /// a bare-name table that cannot tell two `Round`s apart. `None` is
        /// undeclared: the call binds positionally, as every call did before
        /// leaves carried signatures.
        sig: Option<vybe_runtime::component::FuncSig>,
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
        kind: crate::primitives::namespaces::UserGlobalKind,
    },
}

fn terminal(
    forest: &Subtree,
    node: &NamespaceNode,
    path: &str,
    alias_depth: usize,
    fold: crate::primitives::namespaces::Fold,
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
            sig,
            ..
        } => Some(ResolutionTarget::HostCall {
            module: module.clone(),
            func: func.clone(),
            sig: sig.clone(),
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
    use crate::primitives::namespaces::*;

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
        match resolve_path(&["ecma", "json", "stringify"], crate::primitives::namespaces::FOLD_ASCII) {
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
        match resolve_path(&["python", "json", "dumps"], crate::primitives::namespaces::FOLD_ASCII) {
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
        match resolve_path(&["ecma", "json"], crate::primitives::namespaces::FOLD_ASCII) {
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
        match resolve_path(&["python", "js_like", "json", "parse"], crate::primitives::namespaces::FOLD_ASCII) {
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
        assert!(resolve_path(&["a", "x"], crate::primitives::namespaces::FOLD_ASCII).is_none());
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
        assert!(resolve_path(&["ecma", "json", "parse"], crate::primitives::namespaces::FOLD_ASCII).is_some());
        assert!(resolve_path(&["ecma", "math", "max"], crate::primitives::namespaces::FOLD_ASCII).is_some());
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
            resolve_path(&["dotnet", "system", "console"], crate::primitives::namespaces::FOLD_ASCII),
            Some(ResolutionTarget::Ctor { .. })
        ));
        assert!(matches!(
            resolve_path(&["dotnet", "system", "console", "writeline"], crate::primitives::namespaces::FOLD_ASCII),
            Some(ResolutionTarget::HostCall { .. })
        ));
    }

    #[test]
    fn unknown_paths_fall_through() {
        let _g = LOCK.lock().unwrap();
        seed();
        assert!(resolve_path(&["nope"], crate::primitives::namespaces::FOLD_ASCII).is_none());
        assert!(resolve_path(&["ecma", "nope"], crate::primitives::namespaces::FOLD_ASCII).is_none());
        // Leaf with trailing segments is not a namespace path.
        assert!(resolve_path(&["ecma", "json", "stringify", "extra"], crate::primitives::namespaces::FOLD_ASCII).is_none());
    }
}

// ── Registry ───────────────────────────────────────────────────────────────


/// A dot-separated canonical path into the tree (`"ecma.json.stringify"`).
pub type Path = String;

/// Children of a namespace (or the statics of a type).
pub type Subtree = BTreeMap<String, NamespaceNode>;

/// One node in the global namespace forest.
#[derive(Debug, Clone)]
pub enum NamespaceNode {
    /// An interior namespace: `ecma`, `ecma.json`, `dotnet.system`.
    Namespace(Subtree),
    /// One member NAME with several arity-discriminated targets — .NET/Java
    /// style overloads (`Reverse()` vs `Reverse(index, count)`,
    /// `Count()` vs `Count(predicate)`).
    ///
    /// A `Subtree` is keyed by name alone, so without this the last-registered
    /// overload silently wins and every other one resolves to the wrong host
    /// call. Entries are in DECLARATION order and the first matching arity
    /// wins, which is what a descriptor-order scan did before overloads were
    /// tree data.
    Overloads(Vec<(u8, NamespaceNode)>),
    /// A member read and written as a VALUE, with separate targets per
    /// direction. A method has one target; a property has up to two, and they
    /// are frequently different host functions (`controlGetProperty` /
    /// `controlSetProperty`) or different shared emits (`sb_length` /
    /// `sb_set_length`). Either side may be absent (read-only / write-only).
    Property {
        get: Option<Box<NamespaceNode>>,
        set: Option<Box<NamespaceNode>>,
    },
    /// A class/type — CM-typed constructor + static members + instance
    /// method signatures. Instance methods are listed for *metadata* only:
    /// member dispatch is receiver-based (TypeRegistry vtables) and NEVER
    /// resolves through the namespace tree; only `ctor` and `statics` are
    /// reachable by path walk.
    Type {
        ctor: Option<CtorSpec>,
        /// A BACKING constructor call, when the type is constructed by a host
        /// factory or a shared emit rather than by generic field capture
        /// (`ctor`). Held as an ordinary tree node — `Fn` for a host factory,
        /// `CommonEmit` for a shared emit — because that is the vocabulary
        /// every other member already uses; `lookup_type_ctor_target`
        /// translates it at the query boundary. This is what lets a platform
        /// declare `new Dictionary()` as DATA instead of exposing a
        /// `lookup_constructor` hook the compiler has to call into.
        ctor_call: Option<Box<NamespaceNode>>,
        /// Static members, reachable by path walk (`Math.max`).
        statics: Subtree,
        /// INSTANCE members. A `Subtree` — the same node kinds `statics` uses —
        /// so a method resolves to a real target (`Fn { module, func }`,
        /// `CommonEmit`) rather than a bare signature.
        ///
        /// It was `BTreeMap<String, FuncSig>`, which is a SIGNATURE and cannot
        /// carry a target, so nothing could dispatch through it and every
        /// registrar left it empty. That is why the compiler reached into
        /// platform crates for instance-method lookup at all — registering
        /// types was the plugin plan; this field just could not hold the
        /// answer.
        methods: Subtree,
        /// Declared return type per member name, where the platform knows one
        /// (`Object.Equals` → `Boolean`). Lives on the type rather than the
        /// member node so it applies uniformly to `Fn` and `CommonEmit`
        /// members. A platform DECLARES these; the compiler must not carry a
        /// per-platform table of its own.
        member_returns: BTreeMap<String, String>,
    },
    /// A host-backed typed function leaf: resolves to a `CALL_IMPORT` of
    /// `module`/`func` (e.g. `ecma:json` / `stringify`).
    Fn {
        module: String,
        func: String,
        /// The Component Model signature the registrar declared — ONE list of
        /// `(param "name" type)` plus results — or `None` when it declared
        /// nothing, which is what host-export discovery produces.
        ///
        /// Arity is DERIVED from it (`sig.arity()`), never stored beside it.
        /// A registrar that knows only a count declares that count of
        /// unnamed `any` parameters (`host_fn_with_arity`): `any` is the
        /// type's honest spelling of "not stated", not a placeholder. Names
        /// are what let a caller bind an argument by name against this leaf;
        /// a signature with any unnamed parameter binds positionally only.
        sig: Option<vybe_runtime::component::FuncSig>,
        /// A constant argument the registrar bound to this call — a generic
        /// property accessor takes the property NAME as an argument
        /// (`getProperty(this, "Text")`), so the target is the pair
        /// (function, bound name) and not the function alone. Dropping it made
        /// every keyed accessor call the generic getter with no key.
        bound_arg: Option<String>,
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
    /// A declaration made by the COMPILATION UNIT being compiled — a user
    /// class, module or function that lives as a global in the emitted module.
    ///
    /// Every other variant names something the HOST provides: `Fn` is a host
    /// import, `CommonEmit` a shared emit, `Const` a compile-time value, `Type`
    /// a component-model type with a `CtorSpec`. None of them can say "a global
    /// this program declares", which is why user declarations were name-mangled
    /// into flat `HashSet<String>`s instead of living in the tree at all —
    /// `Namespace MyApp.Models` + `Class Customer` became the STRING
    /// `myapp.models.customer` and a namespace was a name prefix, not a node.
    ///
    /// `identity` is that canonical dotted name, unchanged: it is the key every
    /// downstream table still uses (`defined_classes`, `normalized_classes`,
    /// the declared-type hint), so the tree becomes the RESOLVER without
    /// needing all of them migrated at once (namespaceplan.md Phase 6).
    ///
    /// These nodes live only in the per-unit root on `Compiler`, never in the
    /// process-global registry — user declarations are a mount, not tree data.
    /// Every walker in this file reads the global registry, so none of them can
    /// observe one; the compiler's `resolve_segments`, which walks both
    /// forests, matches exhaustively and handles it there.
    UserGlobal {
        identity: Path,
        kind: UserGlobalKind,
        /// Declarations nested UNDER this name. A user declaration is both a
        /// leaf and a container, exactly as `Type` is both a constructable
        /// type and a namespace over its `statics`: `namespace Demo.Sub {
        /// class Demo }` alongside a bare `class Demo` makes `Demo` a type AND
        /// the prefix of `Demo.Sub.Demo`, and both spellings must keep
        /// resolving. Without this the second declaration displaces the first
        /// and whichever arrived last wins.
        children: Subtree,
    },
}

/// What a [`NamespaceNode::UserGlobal`] declares.
///
/// A type and a function are both globals of the unit, but a type position
/// must not accept a function: `Dim c As Repeat` is not a type reference just
/// because `Repeat` is a declared name. Keeping the two apart is what lets one
/// tree answer both questions without either one guessing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserGlobalKind {
    /// A class, structure, enum, interface or module declaration.
    Type,
    /// A function/sub declaration.
    Function,
}

/// Everything the compiler needs to construct a tree `Type` generically,
/// language-neutrally, through the ONE resolver — no per-platform surface.
///
/// A `Type` node carrying this is constructed by: reordering named args by
/// `params` (the shared named-arg machinery), allocating a fresh object,
/// stamping `__type`/`__types` from `ancestry`, and storing each constructor
/// argument into the matching `fields` slot. This is the reusable path that
/// retires the dotnet-surface `lookup_constructor` (namespaceplan.md).
/// The live element type `platforms/web` registers, and therefore the name a
/// platform's [`CtorSpec::ancestry`] must END in for its controls to be
/// allocated with a real rtt rather than the host's one-size element type.
///
/// It lives here, in the tree vocabulary, because it is something a PLATFORM
/// declares — and because the platform crates cannot see `vybe_compiler`
/// (the dependency runs the other way). Matched by name, which is also how a
/// chunk's reserved type slots bind to host `TypeDef`s
/// (`VM::bind_module_type_ids`): declaring this tail gives a control ONE chain
/// answering both `x is Widget` and `x is HTMLElement`, instead of an rtt for
/// one question and a `__types` string array for the other.
pub const DOM_ELEMENT_TYPE: &str = "HTMLElement";

#[derive(Debug, Clone, Default)]
pub struct CtorSpec {
    /// Constructor parameter names, in positional order — drives named-arg
    /// reorder (`Widget(named: v, …)` → positional slots).
    pub params: Vec<String>,
    /// Instance fields to capture from the constructor args, aligned with
    /// `params` (usually identical). Read at the construction site.
    pub fields: Vec<String>,
    /// Type identity chain, self first (`["Scaffold","StatefulWidget",
    /// "Widget"]`) — stamped as `__type` (first) + `__types` (all) so
    /// `is`/`instanceof` matches every ancestor.
    pub ancestry: Vec<String>,
    /// Optional GUI control factory (`new_Panel`/`new_Label`/`new_Button`).
    /// When set, the type is a thin GUI adapter: construction creates the
    /// control and, per arg, either nests a child widget/control or forwards a
    /// scalar property — child vs
    /// scalar detected at runtime. The object IS the control, so widgets
    /// holds all state/layout/rendering. `None` = plain object (no GUI).
    pub control_fn: Option<String>,
    /// How each constructor arg maps onto the control, aligned with `fields`.
    /// Only meaningful when `control_fn` is set.
    pub field_gui: Vec<FieldGui>,
    /// The markup a control is BORN with — its default children.
    ///
    /// The exact companion of the CSS in `ControlElement::declares`, and set at
    /// the same moment for the same reason: it is what the control IS before
    /// any constructor argument is applied, so a program's own writes simply
    /// act on a control that already exists.
    ///
    /// A `BindingNavigator`, a `SplitContainer` and a `MonthCalendar` are each
    /// a container plus fixed chrome — buttons, two panes, a day grid — that
    /// .NET's own constructor materializes (`AddStandardItems`) and a designer
    /// file therefore never writes. With only a tag to declare, they could be
    /// one empty element and nothing more.
    ///
    /// Carried as markup rather than emitted child by child because that is
    /// what the chrome IS — static HTML, readable as HTML in the platform that
    /// declares it. Children built from RUNTIME values are a different job and
    /// keep their adapter (`DataGridView`'s `Columns.Add`).
    ///
    /// The platform owns the vocabulary; the shared path only sets what it was
    /// handed, so no per-language markup lives in a shared crate. Engine-blind:
    /// `SetInnerHtml` is parsed before any engine is entered and re-enters as
    /// ordinary per-tag operations, so every engine builds the same subtree.
    ///
    /// ⛔ Static, so it carries no `id` — an `id` must be unique per document
    /// (DOM §4.9) and two navigators would collide and break `getElementById`.
    /// The parts are addressed by CLASS, which is as targetable and stays valid
    /// however many instances a form holds.
    pub inner_html: Option<String>,
    /// Guest function that answers **"which node does this value contribute?"**
    /// for a value being nested (`NestOrProp`/`Children`), or `None` when the
    /// value nested IS already the node.
    ///
    /// A WinForms control or a GCL widget is its element the moment it is
    /// constructed, so nothing needs asking. Flutter is the framework where
    /// that is not true: `StatelessWidget`/`StatefulWidget` are *configuration*
    /// — `CalculatorPage()` is a description, and the element only exists once
    /// `build()` has run. Nesting the description appends nothing and reports
    /// nothing, which is exactly the blank form.
    ///
    /// The shared path cannot know that, and must not learn it: inflation is
    /// `createState`/`build` with per-widget State, which is Flutter's model and
    /// no one else's. So the platform names the function and the shared path
    /// calls it — one call before the "is this an element" test, identity for
    /// anything already concrete.
    pub nest_coerce: Option<String>,
    /// True for immutable value types whose `==` is by VALUE, not identity
    /// (Flutter `ValueKey`/`Color`/`Offset` override `operator ==`). The
    /// construction site stamps `__value_eq` so the language equality path
    /// compares such instances structurally (by `__type` + fields) instead of
    /// by reference. `false` (default) = reference identity, like a plain class.
    pub value_equality: bool,
}

/// How a GUI-adapter constructor arg is applied to its control.
#[derive(Debug, Clone)]
pub enum FieldGui {
    /// Nest as a child control if the value is a widget, else set it as a
    /// scalar property under this key (`Text.data`→`"Text"`, `Scaffold.body`).
    NestOrProp(String),
    /// The value is a LIST of child widgets (`Column.children`).
    Children,
    /// The value is an event handler wired under this event (`onPressed`
    /// →`"Click"`).
    Event(String),
    /// The value is a child widget whose text IS this control's caption
    /// (`ElevatedButton(child: Text('7'))` → the button's `Text`), rather than
    /// a nested control.
    Caption,
}

impl Default for FieldGui {
    fn default() -> Self {
        FieldGui::NestOrProp(String::new())
    }
}

/// Registry state. The host-export mount flag lives HERE, not in a separate
/// static: it is a validity marker for the tree's contents, so clearing the
/// tree without clearing it leaves a stale "already mounted" claim and the next
/// resolve walks an empty registry. Holding both under one lock makes that
/// impossible to get wrong rather than something every caller must remember.
#[derive(Default)]
pub struct RegistryState {
    pub tree: Subtree,
    host_exports_mounted: bool,
}

static REGISTRY: OnceLock<RwLock<RegistryState>> = OnceLock::new();

/// Read access to the registry for the resolver. The walk lives above this
/// crate (it needs `CapabilityContext`), so it borrows the tree through here
/// rather than the registry owning resolution.
pub fn registry_read() -> std::sync::RwLockReadGuard<'static, RegistryState> {
    registry().read().unwrap()
}

fn registry() -> &'static RwLock<RegistryState> {
    REGISTRY.get_or_init(|| RwLock::new(RegistryState::default()))
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
/// Keys are stored as the REGISTRAR WROTE THEM.
///
/// ⛔ This used to `debug_assert` that the root was lowercase, and that
/// invariant is gone on purpose. It was a shortcut taken while several
/// namespaces and resolvers were being unified into one — a single
/// lowercase-canonical key space was the cheapest way to make one resolver
/// answer for all of them — and it was never a design decision. Lookups now
/// match EXACT first and fold only on a miss (`fold_get`), so a registrar can
/// author `StringBuffer` and a case-insensitive caller still finds it.
///
/// The cost of the old rule was being paid twice already:
/// `languages/python/src/emitter/tree_register.rs` had to BYPASS the
/// `namespace()` helper to keep `OrderedDict` spelled correctly, while
/// `languages/dart/src/tree_register.rs` lowercases despite Dart declaring
/// `case_sensitive = true`. Two registrars, two conventions, and nothing in the
/// tree able to say which one a subtree followed.
///
/// See `documentation/casesensitivityplan.md`.
pub fn register_namespace_tree(root: &str, tree: NamespaceNode) {
    let mut guard = registry().write().unwrap();
    merge_into(&mut guard.tree, root.to_string(), tree);
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
                ctor_call,
                mut statics,
                methods,
                member_returns,
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
                ctor_call,
                statics,
                methods,
                member_returns,
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

/// True when `root` has been registered — cheap membership probe so
/// callers can distinguish "unknown root" from "known root, bad path".
pub fn has_root(root: &str) -> bool {
    registry().read().unwrap().tree.contains_key(root)
}

// ── Type lookup by NAME ─────────────────────────────────────────────────
//
// The tree is path-addressed (`dotnet.system.console`), but the compiler asks
// by receiver TYPE NAME (`Console`, `System.Console`). These walk the tree to
// answer that, so a platform registers its types ONCE — as `Type` nodes — and
// every question is answered from that single registration.
//
// This replaces per-platform lookup hooks on `PlatformDef`, which were a second
// mechanism for data the tree already holds.

/// Find a registered `Type` node by class name, searching only the roots in
/// `scope`. Matches the last path segment, case-insensitively, and accepts a
/// dotted suffix (`System.Console` matches `dotnet.system.console`).
///
/// The scope is what keeps a language from resolving another platform's
/// classes: `Text` is a Flutter widget AND a .NET-ish name, and an unscoped
/// search answers whichever registered first. A language declares which roots
/// its type names come from — a property of the language, not its identity.
fn find_type_node(scope: &[String], class_name: &str, fold: Fold) -> Option<(Subtree, Subtree)> {
    let wanted = class_name.trim().to_string();
    let leaf = wanted.rsplit('.').next().unwrap_or(&wanted).to_string();
    let guard = registry().read().unwrap();

    fn walk(
        node: &NamespaceNode,
        leaf: &str,
        wanted: &str,
        path: &str,
        out: &mut Option<(Subtree, Subtree)>,
        fold: Fold,
    ) {
        if out.is_some() {
            return;
        }
        match node {
            NamespaceNode::Namespace(children) => {
                for (k, v) in children {
                    let next = if path.is_empty() {
                        k.clone()
                    } else {
                        format!("{path}.{k}")
                    };
                    walk(v, leaf, wanted, &next, out, fold);
                }
            }
            NamespaceNode::Type {
                statics, methods, ..
            } => {
                let this_leaf = path.rsplit('.').next().unwrap_or(path);
                if seg_eq(this_leaf, leaf, fold)
                    && (seg_eq(wanted, leaf, fold) || path.to_ascii_lowercase().ends_with(&wanted.to_ascii_lowercase()))
                {
                    *out = Some((statics.clone(), methods.clone()));
                }
            }
            _ => {}
        }
    }

    let mut out = None;
    for root in scope {
        // ⛔ `fold_get`, not `.get`. A scope root comes from a PROFILE
        // (`type_scopes = ["dotnet"]`, lowercase by convention) while a tree
        // root comes from a REGISTRAR, which now writes the declared case.
        // Those two no longer have to agree letter-for-letter, and this is the
        // seam where they meet.
        if let Some(v) = fold_get(&guard.tree, root, fold) {
            walk(v, &leaf, &wanted, root, &mut out, fold);
        }
    }
    out
}

/// Whether a lookup may fall back to a case-insensitive match, and over which
/// alphabet.
///
/// ⛔ `None` means EXACT ONLY. Five languages fold — vb, pascal, cobol, fortran,
/// powershell — plus PHP's callable names; the other twelve must not, because a
/// fold makes them accept programs their real compiler rejects (`math.abs` in
/// Java, `STR(5)` in Python). The policy comes from
/// `Directives::callable_fold()`, so it is the module's DECLARED rule rather
/// than anything this crate infers.
pub type Fold = Option<vybe_ast::CaseAlphabet>;

/// The folding policy as it stood before the fold became conditional: fold
/// every lookup, ASCII.
///
/// Kept for the two callers that must NOT read a module's directives — this
/// file's own unit tests, and the compiler's, which assert the tree's shape
/// rather than any one language's view of it. Production code reads the
/// policy off `Directives::callable_fold()`; a new use here is a bug.
pub const FOLD_ASCII: Fold = Some(vybe_ast::CaseAlphabet::Ascii);

/// Look a key up EXACTLY, then case-insensitively if that misses.
///
/// ⛔ **THE QUERY IS NO LONGER LOWERCASED BEFORE THE LOOKUP.** Every site here
/// used to do `map.get(&name.to_lowercase())`, which can only ever find a
/// lowercase KEY — so the tree was forced to store lowercase-canonical keys,
/// for the benefit of the five case-insensitive languages, at the cost of the
/// twelve case-sensitive ones. That was a shortcut taken while unifying several
/// namespaces and resolvers into one, not a design decision, and it is what
/// `documentation/casesensitivityplan.md` exists to undo.
///
/// Exact-first is what lets both conventions coexist DURING that migration:
/// a registrar that authors `StringBuffer` is found exactly, one that still
/// authors `stringbuffer` is found by the fold, and neither has to change on
/// the same commit as the other. `languages/python/src/emitter/tree_register.rs`
/// already preserves case and had to BYPASS the `namespace()` helper to do it;
/// this is what lets it come back.
///
/// ⛔ AN AMBIGUOUS FOLD IS REFUSED, NOT GUESSED — and it is NOT a declaration
/// bug. Two sibling keys differing only in case are often BOTH LEGAL: `byte` is
/// a C# keyword alias for `System.Byte`, and `platforms/dotnet` registers both
/// deliberately, as it does for `Char`/`char`, `Double`/`double`,
/// `Decimal`/`decimal`. Neither spelling is wrong.
///
/// It costs nothing, because EXACT-FIRST reaches both: C# source writes `byte`
/// and hits `byte`; VB source writes `Byte` and hits `Byte`. The fold is only
/// consulted on a miss, so a legitimate alias pair never needs it.
///
/// When the fold IS reached and two keys answer, picking one would be the
/// `GLOBAL_GET` shape — one key meaning two things depending on who reads it —
/// so this returns `None`, which sends the caller to its next resolution step
/// exactly as a real miss does.
///
/// ⛔ I briefly asserted at `merge_into` that siblings must not collide, on the
/// theory that a collision meant one spelling was wrong. **That rule is false**,
/// it fired on the first legitimate alias pair, and because it was a
/// `debug_assert` it aborted EVERY program on a debug build — blocking every
/// session on the tree. The lookup-time refusal above is the whole of the
/// correct behaviour; there is no registration-time invariant to enforce here.
///
/// ASCII, matching `Scope`'s `eq_ignore_ascii_case`. The tree used to fold with
/// Unicode `to_lowercase()` while scopes folded ASCII — two algorithms in one
/// pipeline, disagreeing on a Turkish dotted `I`, with nothing comparing them.
/// Do two tree path segments name the same thing?
///
/// ⛔ Compares the RAW names rather than pre-lowercasing one side. Both sides
/// used to be lowercased before they met, which forced the KEYS to be
/// lowercase — comparing here instead lets a registrar author `StringBuffer`
/// and a case-insensitive caller still find it.
///
/// ASCII, deliberately: `Scope` folds with `eq_ignore_ascii_case` and this used
/// Unicode `to_lowercase()`, so a Turkish dotted `I` resolved one way as a
/// local and another as a tree path in the same program.
fn seg_eq(a: &str, b: &str, fold: Fold) -> bool {
    match fold {
        None => a == b,
        Some(_) => a.eq_ignore_ascii_case(b),
    }
}

pub fn fold_get<'a>(map: &'a Subtree, key: &str, fold: Fold) -> Option<&'a NamespaceNode> {
    if let Some(hit) = map.get(key) {
        return Some(hit);
    }
    // EXACT ONLY for a language that does not fold.
    fold?;
    let mut found = None;
    for (k, v) in map.iter() {
        if k.eq_ignore_ascii_case(key) {
            if found.is_some() {
                return None;
            }
            found = Some(v);
        }
    }
    found
}

/// [`fold_get`] for the string-valued side tables (`member_returns`).
fn fold_get_str<'a>(
    map: &'a std::collections::BTreeMap<String, String>,
    key: &str,
    fold: Fold,
) -> Option<&'a String> {
    if let Some(hit) = map.get(key) {
        return Some(hit);
    }
    fold?;
    let mut found = None;
    for (k, v) in map.iter() {
        if k.eq_ignore_ascii_case(key) {
            if found.is_some() {
                return None;
            }
            found = Some(v);
        }
    }
    found
}

/// An instance member of `class_name`, from the type's registered `methods`.
pub fn lookup_type_instance_member(
    scope: &[String],
    class_name: &str,
    member: &str,
    fold: Fold,
) -> Option<NamespaceNode> {
    let (_, methods) = find_type_node(scope, class_name, fold)?;
    fold_get(&methods, member, fold).cloned()
}

/// A static member of `class_name`, from the type's registered `statics`.
pub fn lookup_type_static_member(
    scope: &[String],
    class_name: &str,
    member: &str,
    fold: Fold,
) -> Option<NamespaceNode> {
    let (statics, _) = find_type_node(scope, class_name, fold)?;
    fold_get(&statics, member, fold).cloned()
}

/// The declared return type of `member` on `class_name`, if a platform
/// declared one. Replaces a per-platform `static_method_return_type` hook.
pub fn lookup_type_member_return(
    scope: &[String],
    class_name: &str,
    member: &str,
    fold: Fold,
) -> Option<String> {
    let guard = registry().read().unwrap();
    fn walk(node: &NamespaceNode, leaf: &str, member: &str, path: &str, fold: Fold) -> Option<String> {
        match node {
            NamespaceNode::Namespace(children) => children.iter().find_map(|(k, v)| {
                let next = if path.is_empty() {
                    k.clone()
                } else {
                    format!("{path}.{k}")
                };
                walk(v, leaf, member, &next, fold)
            }),
            NamespaceNode::Type { member_returns, .. } => {
                let this_leaf = path.rsplit('.').next().unwrap_or(path);
                seg_eq(this_leaf, leaf, fold)
                    .then(|| fold_get_str(member_returns, member, fold).cloned())
                    .flatten()
            }
            _ => None,
        }
    }
    let leaf = class_name
        .rsplit('.')
        .next()
        .unwrap_or(class_name)
        .to_string();
    scope.iter()
        .find_map(|root| fold_get(&guard.tree, root, fold).and_then(|v| walk(v, &leaf, member, root, fold)))
}

/// True when a type registered UNDER `scope` declares `member` at `arity`.
///
/// Replaces a per-platform `runtime_collection_dispatch_arity` hook: the tree
/// already records each member's arity, and the scope is what makes the answer
/// mean "is this a runtime-dispatched collection member" rather than "does any
/// type anywhere happen to have this name". An unscoped walk answers yes for
/// unrelated types (a GUI control's `Add`) and diverts calls that should never
/// have reached runtime dispatch.
pub fn scope_declares_member_arity(scope: &[&str], member: &str, arity: u8, fold: Fold) -> bool {
    let guard = registry().read().unwrap();
    fn walk(node: &NamespaceNode, member: &str, arity: u8, fold: Fold) -> bool {
        match node {
            NamespaceNode::Namespace(children) => children.values().any(|v| walk(v, member, arity, fold)),
            // ⛔ `fold_get`, not `.get`. The query is no longer lowercased before
            // it arrives, so an exact lookup here misses every key a registrar
            // wrote in a different case — which is what a plain `.get` did the
            // moment the caller stopped folding for it.
            NamespaceNode::Type { methods, .. } => fold_get(methods, member, fold)
                .and_then(|declared| select_overload(declared, arity))
                .is_some_and(
                    |n| matches!(n, NamespaceNode::Fn { sig: Some(s), .. } if s.arity() == arity),
                ),
            _ => false,
        }
    }
    if scope.is_empty() {
        return false;
    }
    let m = member.to_string();
    let mut node = match guard.tree.get(scope[0]) {
        Some(n) => n,
        None => return false,
    };
    for seg in &scope[1..] {
        node = match node {
            // A scope segment is profile data; a child key is registrar data.
            NamespaceNode::Namespace(children) => match fold_get(children, seg, fold) {
                Some(n) => n,
                None => return false,
            },
            _ => return false,
        };
    }
    walk(node, &m, arity, fold)
}

/// An instance METHOD target for `class_name`, from the type's registered
/// members. Returns the descriptor-shaped target the compiler emits, built
/// from the tree node a platform declared — no per-platform hook.
pub fn lookup_type_instance_target(
    scope: &[String],
    class_name: &str,
    member: &str,
    argc: u8,
    fold: Fold,
) -> Option<crate::component_classes::InstanceMethodTarget> {
    let declared = lookup_type_instance_member(scope, class_name, member, fold)?;
    // A zero-arg call on a property IS its getter — `sw.Elapsed` reaches here
    // as a member read with no arguments, and the property node holds the
    // only target there is.
    let declared = match (&declared, argc) {
        (NamespaceNode::Property { get, .. }, 0) => *(get.clone()?),
        _ => declared,
    };
    match select_overload(&declared, argc)?.clone() {
        NamespaceNode::Fn {
            module,
            func,
            sig,
            ..
        } => Some(crate::component_classes::InstanceMethodTarget::Host {
            module,
            func,
            arity: sig.as_ref().map(|s| s.arity()).unwrap_or(argc),
        }),
        NamespaceNode::CommonEmit(emit) => {
            Some(crate::component_classes::InstanceMethodTarget::Common { emit, arity: argc })
        }
        _ => None,
    }
}

// ── Inheritance closure, for REGISTRARS ─────────────────────────────────
//
// `lookup_type_instance_member` is a flat map get, deliberately: a namespace
// tree resolves PATHS, and class inheritance is a different relation. The
// adapter owns that relation because only the adapter holds its rows' parent
// links, and it expands the chain HERE, at registration, so a class's node
// carries its whole inherited surface.
//
// What every adapter needs to do that is the same walk — self first, nearest
// declaration first, case-insensitive, and refusing to spin on a cyclic
// `parent`. It was written out once per derivation instead: four times in the
// dotnet registrar alone (flattened methods, flattened properties, "descends
// from Control", the `CtorSpec` ancestry), three in plib. These two functions
// are that walk, so a registrar supplies its rows' parent link and folds over
// the result.
//
// Deliberately closure-shaped rather than typed over a row struct: adapters
// declare their own row types (`DotnetClass`, `GclClass`, `FlutterClass`) and
// none of them belong in this crate.

/// The inheritance chain of `name`, SELF FIRST then each ancestor, as declared
/// by `parent_of`.
///
/// Nearest first is the contract: fold with a first-wins rule (`or_insert`, or
/// an arity guard on an overload bucket) and an override shadows the base
/// declaration it re-declares, which is what real virtual dispatch does.
///
/// `parent_of` answers the declared parent of a class name, or `None` at the
/// root — and `None` for a name it does not know, which ends the walk. A
/// `parent` chain that cycles stops at the repeat rather than spinning.
///
/// **`name` is ALWAYS the first element**, including when `parent_of` does not
/// know it: the chain is an identity chain, and a class is always itself. A
/// registrar that must distinguish "unregistered" from "root" has to ask its
/// own rows — this walk cannot, and answering `[]` here would silently drop
/// the `__type` stamp of any class whose parent lookup is scoped differently
/// from its member lookup.
pub fn ancestry_of<F>(name: &str, parent_of: F) -> Vec<String>
where
    F: Fn(&str) -> Option<String>,
{
    let mut chain: Vec<String> = Vec::new();
    let mut current = Some(name.to_string());
    while let Some(class) = current {
        if chain.iter().any(|seen| seen.eq_ignore_ascii_case(&class)) {
            break; // cyclic parent chain — refuse to spin
        }
        current = parent_of(&class);
        chain.push(class);
    }
    chain
}

/// True when `name` IS `ancestor` or descends from it, per `parent_of`.
///
/// The membership question a registrar asks to decide whether a row gets a
/// role at all — "is this class a `Control`", "is this a `Throwable`" — which
/// is the same walk as [`ancestry_of`] and was hand-rolled separately every
/// time it was needed.
pub fn inherits<F>(name: &str, ancestor: &str, parent_of: F) -> bool
where
    F: Fn(&str) -> Option<String>,
{
    ancestry_of(name, parent_of)
        .iter()
        .any(|c| c.eq_ignore_ascii_case(ancestor))
}

/// Bundle arity-discriminated targets for one member name. A single entry
/// stays a plain node — overloading is the exception, not the shape.
pub fn overloads(mut entries: Vec<(u8, NamespaceNode)>) -> NamespaceNode {
    if entries.len() == 1 {
        return entries.remove(0).1;
    }
    NamespaceNode::Overloads(entries)
}

/// Select the target for `argc` from a node that may be an overload set.
/// A non-overloaded node answers for every arity — the declaring platform
/// said this name has exactly one target.
pub fn select_overload(node: &NamespaceNode, argc: u8) -> Option<&NamespaceNode> {
    match node {
        NamespaceNode::Overloads(entries) => {
            entries.iter().find(|(a, _)| *a == argc).map(|(_, n)| n)
        }
        other => Some(other),
    }
}

/// The backing constructor NODE a platform declared on `class_name`'s `Type`.
/// Same name matching as `find_type_node` (leaf-wise, case-insensitive,
/// dotted-suffix tolerant) — it reads a different field, so it cannot share
/// that walk.
fn find_type_ctor_call(scope: &[String], class_name: &str, fold: Fold) -> Option<NamespaceNode> {
    let wanted = class_name.trim().to_string();
    let leaf = wanted.rsplit('.').next().unwrap_or(&wanted).to_string();
    let guard = registry().read().unwrap();

    fn walk(node: &NamespaceNode, leaf: &str, wanted: &str, path: &str, fold: Fold) -> Option<NamespaceNode> {
        match node {
            NamespaceNode::Namespace(children) => children.iter().find_map(|(k, v)| {
                let next = if path.is_empty() {
                    k.clone()
                } else {
                    format!("{path}.{k}")
                };
                walk(v, leaf, wanted, &next, fold)
            }),
            NamespaceNode::Type { ctor_call, .. } => {
                let this_leaf = path.rsplit('.').next().unwrap_or(path);
                (seg_eq(this_leaf, leaf, fold)
                    && (seg_eq(wanted, leaf, fold)
                        || path
                            .to_ascii_lowercase()
                            .ends_with(&wanted.to_ascii_lowercase())))
                    .then(|| ctor_call.as_deref().cloned())
                    .flatten()
            }
            _ => None,
        }
    }

    scope.iter()
        .find_map(|root| fold_get(&guard.tree, root, fold).and_then(|v| walk(v, &leaf, &wanted, root, fold)))
}

/// The construction SPEC a platform declared for `class_name`, if the name is a
/// registered `Type` under `scope`. This is what makes a platform base class
/// foldable: the spec is the whole contribution (fields, control factory, GUI
/// field mapping, ancestry).
pub fn lookup_type_ctor_spec(scope: &[String], class_name: &str, fold: Fold) -> Option<CtorSpec> {
    let bare = class_name
        .split(['<', '('])
        .next()
        .unwrap_or(class_name)
        .trim();
    find_type_spec(scope, bare, fold)
}

fn find_type_spec(scope: &[String], class_name: &str, fold: Fold) -> Option<CtorSpec> {
    let wanted = class_name.trim().to_string();
    let leaf = wanted.rsplit('.').next().unwrap_or(&wanted).to_string();
    let guard = registry().read().unwrap();

    fn walk(
        node: &NamespaceNode,
        leaf: &str,
        wanted: &str,
        path: &str,
        fold: Fold,
    ) -> Option<CtorSpec> {
        match node {
            NamespaceNode::Namespace(children) => children.iter().find_map(|(k, v)| {
                let next = if path.is_empty() {
                    k.clone()
                } else {
                    format!("{path}.{k}")
                };
                walk(v, leaf, wanted, &next, fold)
            }),
            NamespaceNode::Type { ctor, .. } => {
                let this_leaf = path.rsplit('.').next().unwrap_or(path);
                (seg_eq(this_leaf, leaf, fold)
                    && (seg_eq(wanted, leaf, fold)
                        || path
                            .to_ascii_lowercase()
                            .ends_with(&wanted.to_ascii_lowercase())))
                    .then(|| ctor.clone())
                    .flatten()
            }
            _ => None,
        }
    }

    scope.iter()
        .find_map(|root| fold_get(&guard.tree, root, fold).and_then(|v| walk(v, &leaf, &wanted, root, fold)))
}

/// The CONSTRUCTOR target for `class_name`, from its registered `Type` node.
///
/// A platform declares a backing constructor as a tree NODE (`Fn` for a host
/// factory, `CommonEmit` for a shared emit) — the same vocabulary every other
/// member uses. This translates that node into the descriptor-shaped target
/// the compiler emits, so there is no per-platform `lookup_constructor` hook.
pub fn lookup_type_ctor_target(
    scope: &[String],
    class_name: &str,
    fold: Fold,
) -> Option<crate::component_classes::ConstructorTarget> {
    use crate::component_classes::{ConstructorTarget, HostTarget};
    // `Dictionary<K, V>` / `List(Of T)` name the same registered type as the
    // bare name — strip the generic argument list before matching.
    let bare = class_name
        .split(['<', '('])
        .next()
        .unwrap_or(class_name)
        .trim();
    match find_type_ctor_call(scope, bare, fold)? {
        NamespaceNode::Fn { module, func, .. } => {
            Some(ConstructorTarget::Host(HostTarget { module, name: func }))
        }
        NamespaceNode::CommonEmit(emit) => Some(ConstructorTarget::Common(emit)),
        _ => None,
    }
}

/// An instance PROPERTY target for `class_name`.
pub fn lookup_type_property_target(
    scope: &[String],
    class_name: &str,
    member: &str,
    fold: Fold,
) -> Option<crate::component_classes::InstancePropertyTarget> {
    property_target(
        lookup_type_instance_member(scope, class_name, member, fold)?,
        false,
    )
}

/// The SETTER target for `class_name.member`. A property's two directions are
/// different targets, so a read-only lookup cannot answer for a write.
pub fn lookup_type_property_setter_target(
    scope: &[String],
    class_name: &str,
    member: &str,
    fold: Fold,
) -> Option<crate::component_classes::InstancePropertyTarget> {
    property_target(
        lookup_type_instance_member(scope, class_name, member, fold)?,
        true,
    )
}

fn property_target(
    declared: NamespaceNode,
    want_setter: bool,
) -> Option<crate::component_classes::InstancePropertyTarget> {
    use crate::component_classes::InstancePropertyTarget;
    let node = match declared {
        NamespaceNode::Property { get, set } => *(if want_setter { set } else { get })?,
        // A plain leaf answers reads only: a method-shaped member has no
        // write direction to route a store through.
        other if !want_setter => other,
        _ => return None,
    };
    match node {
        NamespaceNode::Fn {
            module,
            func,
            bound_arg,
            ..
        } => Some(InstancePropertyTarget::Host {
            module,
            func,
            key: bound_arg,
        }),
        NamespaceNode::CommonEmit(emit) => Some(InstancePropertyTarget::Common { emit }),
        _ => None,
    }
}

/// Does the tree DECLARE this dotted path?
///
/// The general "is this name on that module's surface" question, answered by
/// the one resolver instead of by a per-language table. `hasattr(logging,
/// "FileHandler")`, `from logging import FileHandler` and
/// `logging.FileHandler` are three spellings of it, and a language that keeps
/// its own static surface list has three answers that can disagree — which is
/// exactly what happened: python's walker-side `py_module_surface` said the
/// name was absent while the tree resolved it fine.
///
/// A `Type` node is a leaf for this purpose: `python.ssl.SSLContext` exists,
/// and whether `SSLContext` has some member is a different query
/// (`lookup_type_instance_member`).
pub fn declares_path(path: &[&str], fold: Fold) -> bool {
    if path.is_empty() {
        return false;
    }
    let guard = registry().read().unwrap();
    let Some(mut node) = guard.tree.get(path[0]) else {
        return false;
    };
    for seg in &path[1..] {
        node = match node {
            NamespaceNode::Namespace(children) => match fold_get(children, seg, fold) {
                Some(n) => n,
                None => return false,
            },
            // Statics hang off a type, so `Math.PI` keeps resolving.
            NamespaceNode::Type { statics, .. } => match fold_get(statics, seg, fold) {
                Some(n) => n,
                None => return false,
            },
            _ => return false,
        };
    }
    true
}

/// True when any registered platform contributed a `Type` node for this name.
pub fn is_registered_type(scope: &[String], class_name: &str, fold: Fold) -> bool {
    find_type_node(scope, class_name, fold).is_some()
}

/// Test-only: wipe the registry so unit tests are order-independent.
#[doc(hidden)]
pub fn clear_registry_for_tests() {
    // One assignment resets tree AND mount flag together — they cannot drift.
    *registry().write().unwrap() = RegistryState::default();
}

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
/// Keys keep the case the host export declared, as does the leaf payload
/// (`isArray`) — emission has always used the payload, which is why the
/// `matchAll` bug class stayed fixed even while KEYS were folded. Now neither
/// side is folded and `fold_get` handles a case-insensitive caller at lookup.

// ── Construction helpers ────────────────────────────────────────────────

/// An untyped host-fn leaf (`Any` params/results) — the common case until
/// platforms register real CM signatures.
pub fn host_fn(module: &str, func: &str) -> NamespaceNode {
    NamespaceNode::Fn {
        module: module.to_string(),
        func: func.to_string(),
        sig: None,
        bound_arg: None,
    }
}

/// A host-backed leaf that binds a constant argument — a generic property
/// accessor, which takes the property name (`getProperty(this, "Text")`).
/// The pair IS the target; the function alone is not.
pub fn host_fn_keyed(module: &str, func: &str, key: &str) -> NamespaceNode {
    NamespaceNode::Fn {
        module: module.to_string(),
        func: func.to_string(),
        sig: None,
        bound_arg: Some(key.to_string()),
    }
}

/// A property member with per-direction targets. Either side may be absent.
pub fn property(get: Option<NamespaceNode>, set: Option<NamespaceNode>) -> NamespaceNode {
    NamespaceNode::Property {
        get: get.map(Box::new),
        set: set.map(Box::new),
    }
}

/// A host-backed leaf whose arity the registrar knows: `arity` unnamed
/// `any` parameters and an `any` result — the count is stated, nothing else.
pub fn host_fn_with_arity(module: &str, func: &str, arity: u8) -> NamespaceNode {
    use vybe_runtime::component::{FuncSig, Param, ValType};
    host_fn_with_sig(
        module,
        FuncSig {
            name: func.to_string(),
            params: Param::unnamed_list(vec![ValType::Any; arity as usize]),
            results: vec![ValType::Any],
        },
    )
}

/// A host-backed leaf with its declared Component Model signature. The
/// leaf's function name is `sig.name`. Parameter names describe the declared
/// parameters only — a leaf naming its receiver (`this`/`self`) would be
/// read as taking one more argument than every caller passes, so it is
/// refused here, at registration, where the registrar can see it.
pub fn host_fn_with_sig(module: &str, sig: vybe_runtime::component::FuncSig) -> NamespaceNode {
    assert!(
        !sig.params
            .iter()
            .any(|p| p.name.eq_ignore_ascii_case("this") || p.name.eq_ignore_ascii_case("self")),
        "{module}/{} names its receiver as a parameter; names describe declared parameters only",
        sig.name
    );
    NamespaceNode::Fn {
        module: module.to_string(),
        func: sig.name.clone(),
        sig: Some(sig),
        bound_arg: None,
    }
}

/// A namespace from a list of children.
pub fn namespace(children: Vec<(&str, NamespaceNode)>) -> NamespaceNode {
    NamespaceNode::Namespace(
        children
            .into_iter()
            // Keys keep the case the registrar wrote. The lowercase-canonical
            // assertion that used to live here is gone — see
            // `register_namespace_tree`.
            .map(|(k, v)| (k.to_string(), v))
            .collect(),
    )
}

pub fn mount_host_exports<I>(exports: I)
where
    I: IntoIterator<Item = (String, String)>,
{
    let mut guard = registry().write().unwrap();
    // Checked and set under the SAME lock as the tree it describes.
    if guard.host_exports_mounted {
        return;
    }
    guard.host_exports_mounted = true;
    for (module, func) in exports {
        // "ecma:json" → [ecma, json]; "wasi:cli/stdout" → [wasi, cli, stdout]
        let mut segments: Vec<String> = Vec::new();
        let mut rest = module.as_str();
        // Segments keep the case the host export declared. Host module names
        // ("ecma:json", "wasi:cli/stdout") are lowercase already, so this is
        // not a behaviour change today — it is the invariant changing, so a
        // host that later exports a cased name is not silently folded.
        if let Some((pkg, tail)) = rest.split_once(':') {
            segments.push(pkg.to_string());
            rest = tail;
        }
        for part in rest.split('/') {
            if !part.is_empty() {
                segments.push(part.to_string());
            }
        }
        if segments.is_empty() {
            continue;
        }
        segments.push(func.to_string());

        // Build the nested single-child tree for this leaf, then merge.
        // Host exports are discovered from the function registry, which does
        // not report arity —  states that honestly.
        let mut node = NamespaceNode::Fn {
            module: module.clone(),
            func: func.clone(),
            sig: None,
            bound_arg: None,
        };
        while segments.len() > 1 {
            let key = segments.pop().unwrap();
            let mut children = Subtree::new();
            children.insert(key, node);
            node = NamespaceNode::Namespace(children);
        }
        merge_into(&mut guard.tree, segments.pop().unwrap(), node);
    }
}
