//! Namespace-tree REGISTRY — the shared contract between plugins and the
//! compiler.
//!
//! Lives in `vybe_runtime` because that is the crate EVERY writer already
//! depends on: all twelve languages, every platform, `vybe_compiler::emitter` and
//! `vybe_compiler`. That is what lets a plugin register its surface from
//! `Plugin::init` with no dependency edge on the compiler — and it is what a
//! `dlopen`'d dylib needs, since the registry must live in the crate the host
//! and the plugin both link.
//!
//! It was in `vybe_compiler::emitter`, which sits ABOVE the platform crates, so a
//! platform could not be registered from a shared plugin list without a cycle
//! (`vybe_compiler::emitter -> platform_dotnet -> vybe_compiler::emitter`).
//!
//! WRITE side (plugins): `register_namespace_tree`, plus the node constructors
//! `host_fn` / `namespace`. Verified usage: platforms and languages only ever
//! write — neither calls `resolve_path`.
//!
//! READ side (compiler): `resolve_path` is a raw walk. The POLICY around it —
//! ambient roots, alias resolution, `Resolution` — is compiler behaviour and
//! stays in `vybe_compiler::primitives::resolver`.

use std::collections::BTreeMap;
use std::sync::{OnceLock, RwLock};

use crate::Value;

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
        set: Option<Box<NamespaceNode>> },
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
        member_returns: BTreeMap<String, String> },
    /// A host-backed typed function leaf: resolves to a `CALL_IMPORT` of
    /// `module`/`func` (e.g. `ecma:json` / `stringify`).
    Fn {
        module: String,
        func: String,
        /// User-visible arity when the registrar knows it (a descriptor's
        /// `MethodDef.arity`), else `None`.
        ///
        /// This replaced a `sig: FuncSig`, which every consumer discarded
        /// (`HostCall { module, func, .. }`) and which could not express what
        /// registrars actually know: descriptors carry an ARITY, not param
        /// types. Encoding a count as placeholder `ValType`s would be inventing
        /// type information — the same mistake as a defaulted `CtorSpec`.
        arity: Option<u8>,
        /// A constant argument the registrar bound to this call — the generic
        /// `vybe:gui` accessors take the property NAME as an argument
        /// (`controlGetProperty(this, "Text")`), so the target is the pair
        /// (function, bound name) and not the function alone. Dropping it made
        /// every keyed accessor call the generic getter with no key.
        bound_arg: Option<String> },
    /// A `common:<cat>.<op>` dispatch leaf — emitted through the shared
    /// common-emit dispatcher rather than a direct host call.
    CommonEmit(String),
    /// A compile-time constant value (`Math.PI`-class leaves).
    Const(Value),
    /// This name → a leaf living at a different canonical path. The
    /// source≠canonical reconciliation point (`python.json.dumps` →
    /// `Alias("ecma.json.stringify")`).
    Alias(Path) }

/// Everything the compiler needs to construct a tree `Type` generically,
/// language-neutrally, through the ONE resolver — no per-platform surface.
///
/// A `Type` node carrying this is constructed by: reordering named args by
/// `params` (the shared named-arg machinery), allocating a fresh object,
/// stamping `__type`/`__types` from `ancestry`, and storing each constructor
/// argument into the matching `fields` slot. This is the reusable path that
/// retires the dotnet-surface `lookup_constructor` (namespaceplan.md).
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
    /// `vybe:gui` control and, per arg, either nests a child widget/control
    /// (`controlsAdd`) or forwards a scalar (`controlSetProperty`) — child vs
    /// scalar detected at runtime. The object IS the control, so vybe_widgets
    /// holds all state/layout/rendering. `None` = plain object (no GUI).
    pub control_fn: Option<String>,
    /// How each constructor arg maps onto the control, aligned with `fields`.
    /// Only meaningful when `control_fn` is set.
    pub field_gui: Vec<FieldGui>,
    /// True for immutable value types whose `==` is by VALUE, not identity
    /// (Flutter `ValueKey`/`Color`/`Offset` override `operator ==`). The
    /// construction site stamps `__value_eq` so the language equality path
    /// compares such instances structurally (by `__type` + fields) instead of
    /// by reference. `false` (default) = reference identity, like a plain class.
    pub value_equality: bool }

/// How a GUI-adapter constructor arg is applied to its `vybe:gui` control.
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
    Caption }

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
    host_exports_mounted: bool }

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
/// Keys must already be lowercase-canonical — see the module docs.
pub fn register_namespace_tree(root: &str, tree: NamespaceNode) {
    debug_assert_eq!(
        root,
        root.to_lowercase(),
        "namespace tree keys are stored lowercase-canonical; canon at the boundary"
    );
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
                member_returns },
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
                member_returns };
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
fn find_type_node(scope: &[String], class_name: &str) -> Option<(Subtree, Subtree)> {
    let wanted = class_name.trim().to_lowercase();
    let leaf = wanted.rsplit('.').next().unwrap_or(&wanted).to_string();
    let guard = registry().read().unwrap();

    fn walk(
        node: &NamespaceNode,
        leaf: &str,
        wanted: &str,
        path: &str,
        out: &mut Option<(Subtree, Subtree)>,
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
                    walk(v, leaf, wanted, &next, out);
                }
            }
            NamespaceNode::Type {
                statics, methods, ..
            } => {
                let this_leaf = path.rsplit('.').next().unwrap_or(path);
                if this_leaf == leaf && (wanted == leaf || path.ends_with(wanted)) {
                    *out = Some((statics.clone(), methods.clone()));
                }
            }
            _ => {}
        }
    }

    let mut out = None;
    for root in scope {
        if let Some(v) = guard.tree.get(root) {
            walk(v, &leaf, &wanted, root, &mut out);
        }
    }
    out
}

/// An instance member of `class_name`, from the type's registered `methods`.
pub fn lookup_type_instance_member(
    scope: &[String],
    class_name: &str,
    member: &str,
) -> Option<NamespaceNode> {
    let (_, methods) = find_type_node(scope, class_name)?;
    methods.get(&member.to_lowercase()).cloned()
}

/// A static member of `class_name`, from the type's registered `statics`.
pub fn lookup_type_static_member(
    scope: &[String],
    class_name: &str,
    member: &str,
) -> Option<NamespaceNode> {
    let (statics, _) = find_type_node(scope, class_name)?;
    statics.get(&member.to_lowercase()).cloned()
}

/// The declared return type of `member` on `class_name`, if a platform
/// declared one. Replaces a per-platform `static_method_return_type` hook.
pub fn lookup_type_member_return(
    scope: &[String],
    class_name: &str,
    member: &str,
) -> Option<String> {
    let guard = registry().read().unwrap();
    fn walk(node: &NamespaceNode, leaf: &str, member: &str, path: &str) -> Option<String> {
        match node {
            NamespaceNode::Namespace(children) => children.iter().find_map(|(k, v)| {
                let next = if path.is_empty() { k.clone() } else { format!("{path}.{k}") };
                walk(v, leaf, member, &next)
            }),
            NamespaceNode::Type { member_returns, .. } => {
                let this_leaf = path.rsplit('.').next().unwrap_or(path);
                (this_leaf == leaf)
                    .then(|| member_returns.get(&member.to_lowercase()).cloned())
                    .flatten()
            }
            _ => None }
    }
    let leaf = class_name.rsplit('.').next().unwrap_or(class_name).to_lowercase();
    scope
        .iter()
        .find_map(|root| guard.tree.get(root).and_then(|v| walk(v, &leaf, member, root)))
}

/// True when a type registered UNDER `scope` declares `member` at `arity`.
///
/// Replaces a per-platform `runtime_collection_dispatch_arity` hook: the tree
/// already records each member's arity, and the scope is what makes the answer
/// mean "is this a runtime-dispatched collection member" rather than "does any
/// type anywhere happen to have this name". An unscoped walk answers yes for
/// unrelated types (a GUI control's `Add`) and diverts calls that should never
/// have reached runtime dispatch.
pub fn scope_declares_member_arity(scope: &[&str], member: &str, arity: u8) -> bool {
    let guard = registry().read().unwrap();
    fn walk(node: &NamespaceNode, member: &str, arity: u8) -> bool {
        match node {
            NamespaceNode::Namespace(children) => {
                children.values().any(|v| walk(v, member, arity))
            }
            NamespaceNode::Type { methods, .. } => methods
                .get(member)
                .and_then(|declared| select_overload(declared, arity))
                .is_some_and(
                    |n| matches!(n, NamespaceNode::Fn { arity: Some(a), .. } if *a == arity),
                ),
            _ => false }
    }
    if scope.is_empty() {
        return false;
    }
    let m = member.to_lowercase();
    let mut node = match guard.tree.get(scope[0]) {
        Some(n) => n,
        None => return false };
    for seg in &scope[1..] {
        node = match node {
            NamespaceNode::Namespace(children) => match children.get(*seg) {
                Some(n) => n,
                None => return false },
            _ => return false };
    }
    walk(node, &m, arity)
}

/// An instance METHOD target for `class_name`, from the type's registered
/// members. Returns the descriptor-shaped target the compiler emits, built
/// from the tree node a platform declared — no per-platform hook.
pub fn lookup_type_instance_target(
    scope: &[String],
    class_name: &str,
    member: &str,
    argc: u8,
) -> Option<crate::component_model::InstanceMethodTarget> {
    let declared = lookup_type_instance_member(scope, class_name, member)?;
    // A zero-arg call on a property IS its getter — `sw.Elapsed` reaches here
    // as a member read with no arguments, and the property node holds the
    // only target there is.
    let declared = match (&declared, argc) {
        (NamespaceNode::Property { get, .. }, 0) => *(get.clone()?),
        _ => declared };
    match select_overload(&declared, argc)?.clone() {
        NamespaceNode::Fn {
            module,
            func,
            arity,
            ..
        } => Some(crate::component_model::InstanceMethodTarget::Host {
            module,
            func,
            arity: arity.unwrap_or(argc) }),
        NamespaceNode::CommonEmit(emit) => {
            Some(crate::component_model::InstanceMethodTarget::Common { emit, arity: argc })
        }
        _ => None }
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
        NamespaceNode::Overloads(entries) => entries
            .iter()
            .find(|(a, _)| *a == argc)
            .map(|(_, n)| n),
        other => Some(other) }
}

/// The backing constructor NODE a platform declared on `class_name`'s `Type`.
/// Same name matching as `find_type_node` (leaf-wise, case-insensitive,
/// dotted-suffix tolerant) — it reads a different field, so it cannot share
/// that walk.
fn find_type_ctor_call(scope: &[String], class_name: &str) -> Option<NamespaceNode> {
    let wanted = class_name.trim().to_lowercase();
    let leaf = wanted.rsplit('.').next().unwrap_or(&wanted).to_string();
    let guard = registry().read().unwrap();

    fn walk(node: &NamespaceNode, leaf: &str, wanted: &str, path: &str) -> Option<NamespaceNode> {
        match node {
            NamespaceNode::Namespace(children) => children.iter().find_map(|(k, v)| {
                let next = if path.is_empty() {
                    k.clone()
                } else {
                    format!("{path}.{k}")
                };
                walk(v, leaf, wanted, &next)
            }),
            NamespaceNode::Type { ctor_call, .. } => {
                let this_leaf = path.rsplit('.').next().unwrap_or(path);
                (this_leaf == leaf && (wanted == leaf || path.ends_with(wanted)))
                    .then(|| ctor_call.as_deref().cloned())
                    .flatten()
            }
            _ => None }
    }

    scope
        .iter()
        .find_map(|root| guard.tree.get(root).and_then(|v| walk(v, &leaf, &wanted, root)))
}

/// The construction SPEC a platform declared for `class_name`, if the name is a
/// registered `Type` under `scope`. This is what makes a platform base class
/// foldable: the spec is the whole contribution (fields, control factory, GUI
/// field mapping, ancestry).
pub fn lookup_type_ctor_spec(scope: &[String], class_name: &str) -> Option<CtorSpec> {
    let bare = class_name
        .split(['<', '('])
        .next()
        .unwrap_or(class_name)
        .trim();
    find_type_spec(scope, bare)
}

fn find_type_spec(scope: &[String], class_name: &str) -> Option<CtorSpec> {
    let wanted = class_name.trim().to_lowercase();
    let leaf = wanted.rsplit('.').next().unwrap_or(&wanted).to_string();
    let guard = registry().read().unwrap();

    fn walk(node: &NamespaceNode, leaf: &str, wanted: &str, path: &str) -> Option<CtorSpec> {
        match node {
            NamespaceNode::Namespace(children) => children.iter().find_map(|(k, v)| {
                let next = if path.is_empty() {
                    k.clone()
                } else {
                    format!("{path}.{k}")
                };
                walk(v, leaf, wanted, &next)
            }),
            NamespaceNode::Type { ctor, .. } => {
                let this_leaf = path.rsplit('.').next().unwrap_or(path);
                (this_leaf == leaf && (wanted == leaf || path.ends_with(wanted)))
                    .then(|| ctor.clone())
                    .flatten()
            }
            _ => None }
    }

    scope
        .iter()
        .find_map(|root| guard.tree.get(root).and_then(|v| walk(v, &leaf, &wanted, root)))
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
) -> Option<crate::component_model::ConstructorTarget> {
    use crate::component_model::{ConstructorTarget, HostTarget};
    // `Dictionary<K, V>` / `List(Of T)` name the same registered type as the
    // bare name — strip the generic argument list before matching.
    let bare = class_name
        .split(['<', '('])
        .next()
        .unwrap_or(class_name)
        .trim();
    match find_type_ctor_call(scope, bare)? {
        NamespaceNode::Fn { module, func, .. } => Some(ConstructorTarget::Host(HostTarget {
            module,
            name: func })),
        NamespaceNode::CommonEmit(emit) => Some(ConstructorTarget::Common(emit)),
        _ => None }
}

/// An instance PROPERTY target for `class_name`.
pub fn lookup_type_property_target(
    scope: &[String],
    class_name: &str,
    member: &str,
) -> Option<crate::component_model::InstancePropertyTarget> {
    property_target(lookup_type_instance_member(scope, class_name, member)?, false)
}

/// The SETTER target for `class_name.member`. A property's two directions are
/// different targets, so a read-only lookup cannot answer for a write.
pub fn lookup_type_property_setter_target(
    scope: &[String],
    class_name: &str,
    member: &str,
) -> Option<crate::component_model::InstancePropertyTarget> {
    property_target(lookup_type_instance_member(scope, class_name, member)?, true)
}

fn property_target(
    declared: NamespaceNode,
    want_setter: bool,
) -> Option<crate::component_model::InstancePropertyTarget> {
    use crate::component_model::InstancePropertyTarget;
    let node = match declared {
        NamespaceNode::Property { get, set } => {
            *(if want_setter { set } else { get })?
        }
        // A plain leaf answers reads only: a method-shaped member has no
        // write direction to route a store through.
        other if !want_setter => other,
        _ => return None };
    match node {
        NamespaceNode::Fn {
            module,
            func,
            bound_arg,
            ..
        } => Some(InstancePropertyTarget::Host {
            module,
            func,
            key: bound_arg }),
        NamespaceNode::CommonEmit(emit) => Some(InstancePropertyTarget::Common { emit }),
        _ => None }
}

/// True when any registered platform contributed a `Type` node for this name.
pub fn is_registered_type(scope: &[String], class_name: &str) -> bool {
    find_type_node(scope, class_name).is_some()
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
/// Keys are lowercased for storage (tree canon); the leaf payload keeps
/// the host's true casing (`isArray`) so emission never mangles the
/// name — the `matchAll` bug class.

// ── Construction helpers ────────────────────────────────────────────────

/// An untyped host-fn leaf (`Any` params/results) — the common case until
/// platforms register real CM signatures.
pub fn host_fn(module: &str, func: &str) -> NamespaceNode {
    NamespaceNode::Fn {
        module: module.to_string(),
        func: func.to_string(),
        arity: None,
        bound_arg: None }
}

/// A host-backed leaf that binds a constant argument — the generic `vybe:gui`
/// property accessors, which take the property name (`controlGetProperty(this,
/// "Text")`). The pair IS the target; the function alone is not.
pub fn host_fn_keyed(module: &str, func: &str, key: &str) -> NamespaceNode {
    NamespaceNode::Fn {
        module: module.to_string(),
        func: func.to_string(),
        arity: None,
        bound_arg: Some(key.to_string()) }
}

/// A property member with per-direction targets. Either side may be absent.
pub fn property(get: Option<NamespaceNode>, set: Option<NamespaceNode>) -> NamespaceNode {
    NamespaceNode::Property {
        get: get.map(Box::new),
        set: set.map(Box::new) }
}

/// A host-backed leaf whose arity the registrar knows.
pub fn host_fn_with_arity(module: &str, func: &str, arity: u8) -> NamespaceNode {
    NamespaceNode::Fn {
        module: module.to_string(),
        func: func.to_string(),
        arity: Some(arity),
        bound_arg: None }
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
        // Host exports are discovered from the function registry, which does
        // not report arity —  states that honestly.
        let mut node = NamespaceNode::Fn {
            module: module.clone(),
            func: func.clone(),
            arity: None,
            bound_arg: None };
        while segments.len() > 1 {
            let key = segments.pop().unwrap();
            let mut children = Subtree::new();
            children.insert(key, node);
            node = NamespaceNode::Namespace(children);
        }
        merge_into(&mut guard.tree, segments.pop().unwrap(), node);
    }
}
