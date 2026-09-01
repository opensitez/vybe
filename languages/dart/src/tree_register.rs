//! `dart.*` namespace-tree registration (namespaceplan.md Phase 5).
//!
//! Mirrors the dotnet registrar: the Dart language contributes DATA — its
//! profile `[builtins]` table and namespace aliases, the same data its
//! emit dispatch and linker execute — to the shared namespace tree.
//! Resolution LOGIC lives only in the common resolver; any language can
//! walk `dart.math.sqrt` (an alias leaf onto the canonical `ecma.math`).
//!
//! Leaf rules (dotnet template + the plan's alias-leaves):
//! - `common:dart.<fn>` emits register as `CommonEmit` leaves;
//! - host-backed builtins register as `Fn` leaves at their (dotted)
//!   builtin path (`dart.double.parse`);
//! - profile namespace aliases (`math` → `ecma:math`) register as
//!   `Alias` leaves onto the canonical tree path (`dart.math` →
//!   `ecma.math`) — source names reconciled to canonical names as data;
//! - opcode/intrinsic builtins have no process-global target — skipped.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};
use vybe_runtime::profile::{BuiltinEmit, EsmDefault, parse_profile};

#[derive(Clone, Copy)]
enum AdapterCtor {
    Common(&'static str),
    Host(&'static str, &'static str),
}

#[derive(Clone, Copy)]
struct AdapterType {
    library: &'static str,
    name: &'static str,
    ctor: AdapterCtor,
}

/// Dart library classes whose runtime value is backed by an adapter/common
/// emit rather than by an AST class body.
///
/// This is DATA for the namespace tree: construction still goes through the
/// common `ExprKind::New` -> `lookup_type_ctor_target` path, which stamps
/// identity and keeps these reachable as `dart.core.RegExp`,
/// `dart.collection.Queue`, etc. The walker reads the same table only to
/// normalize Dart's constructor syntax (`RegExp(...)`, no `new`) into `New`.
const ADAPTER_TYPES: &[AdapterType] = &[
    AdapterType {
        library: "core",
        name: "RegExp",
        ctor: AdapterCtor::Common("dart.regexp_new"),
    },
    AdapterType {
        library: "core",
        name: "Map",
        ctor: AdapterCtor::Common("dart.map_new"),
    },
    AdapterType {
        library: "core",
        name: "MapEntry",
        ctor: AdapterCtor::Common("dart.map_entry"),
    },
    AdapterType {
        library: "core",
        name: "Expando",
        ctor: AdapterCtor::Common("dart.map_new"),
    },
    AdapterType {
        library: "core",
        name: "Stopwatch",
        ctor: AdapterCtor::Common("dart.stopwatch_new"),
    },
    AdapterType {
        library: "collection",
        name: "Queue",
        ctor: AdapterCtor::Common("dart.stream_empty"),
    },
    AdapterType {
        library: "collection",
        name: "SplayTreeMap",
        ctor: AdapterCtor::Common("dart.sorted_map_new"),
    },
    AdapterType {
        library: "collection",
        name: "SplayTreeSet",
        ctor: AdapterCtor::Common("dart.set_new"),
    },
    AdapterType {
        library: "math",
        name: "Random",
        ctor: AdapterCtor::Host("wasi:random/insecure", "get-insecure-random-u64"),
    },
];

pub(crate) fn is_adapter_type(name: &str) -> bool {
    ADAPTER_TYPES.iter().any(|ty| ty.name == name)
}

fn adapter_ctor_node(ctor: AdapterCtor) -> NamespaceNode {
    match ctor {
        AdapterCtor::Common(name) => NamespaceNode::CommonEmit(name.to_string()),
        AdapterCtor::Host(module, func) => namespaces::host_fn(module, func),
    }
}

/// Insert `node` at the dotted `path` under `root`, creating interior
/// namespaces as needed.
///
/// Keys keep the case the source wrote. ⛔ Dart declares
/// `case_sensitive = true`, so folding its type names here was wrong on its own
/// terms — `StringBuffer` is not `stringbuffer` in Dart. It only worked because
/// every tree lookup lowercased its query too, which is the shortcut
/// `documentation/casesensitivityplan.md` exists to undo. Lookups now match
/// EXACT first and fold only on a miss, so the real spelling is what resolves,
/// and a case-insensitive language reaching the same node still finds it.
fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let leaf = segments.pop().expect("non-empty path");
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return; // leaf/namespace collision: first registration wins
        };
        cursor = children;
    }
    cursor.entry(leaf.to_string()).or_insert(node);
}

/// `"ecma:math"` → `"ecma.math"`, `"wasi:cli/terminal"` → `"wasi.cli.terminal"`.
fn module_tree_path(module: &str) -> String {
    // No fold: host module paths are already lowercase, and folding here would
    // silently rewrite a cased one.
    module.replace([':', '/'], ".")
}

/// `dart:core` TYPES, declared in the namespace tree.
///
/// The tree declares the NAME — that is what makes `dart.core.StringBuffer`
/// reachable through the common resolver from any language, the same way
/// `flutter.*` and `dotnet.*` are. It deliberately carries **no `ctor_call`**:
/// the class itself is an ordinary `ClassDecl` (`core_classes/`) that
/// normalises and compiles like any user class, so construction is already
/// answered by `compile_class`'s rtt and vtable, and a `ctor_call` here would
/// intercept it and put back the anonymous `struct.new 0` this migration
/// removes.
///
/// `methods` carries the PROPERTY leaves (`core_properties`). A type node with
/// no leaves is incomplete — namespaceplan.md §"Leaves" — and a `Property`
/// leaf is read by the member-read path, not by receiver dispatch, so it does
/// not shadow the compiled class's own methods.
///
/// `member_returns` stays, because a declared return type is knowledge the
/// tree owns and the class body does not restate.
///
/// The name list comes from `core_classes::CORE_CLASSES`, so the tree cannot
/// declare a type the AST does not build, or miss one it does.
/// The instance PROPERTY leaves of a core type.
///
/// namespaceplan.md §"Leaves": *"a package or type node alone is incomplete …
/// property getters/setters live under the type's `methods`"*, and the leaf
/// kind for a member read and written as a VALUE is `Property { get, set }` —
/// whose own doc names `sb_length` / `sb_set_length` as the motivating case
/// (`vybe_compiler::primitives::namespaces`). This is that declaration.
///
/// The shared member read consumes it directly (`expressions.rs:3081`): with
/// `type_scopes` non-empty and the receiver's static type hint naming a
/// registered type, `lookup_type_property_target` answers and the leaf's emit
/// runs. So a declared property needs NO `[value_methods]` row and no walker
/// force-call — those are the duplicate mechanism this replaces.
///
/// `get` points at `dart.length`, which is `emit_dart_length` — the consumer
/// that reads `ProtocolSlot::Len` off the receiver before probing its runtime
/// shape. Routing through the tree keeps the ROLE as the resolution; it only
/// changes what carries the read there.
///
/// Keys are lowercase: `lookup_type_instance_member` folds the member name.
fn core_properties(name: &str) -> Subtree {
    let getters: &[(&str, &str)] = match name {
        "StringBuffer" => &[
            ("length", "dart.length"),
            // Emptiness is its OWN slot. The class spells `bool get isEmpty`,
            // which `protocol.rs` maps to `ProtocolSlot::IsEmpty` and
            // `normalize_class.rs` publishes from the GETTER; `dart.is_empty`
            // is now the consumer that asks for it before falling back to
            // `length == 0`. Declaring the leaf is what lets the member READ
            // reach that consumer without the walker forging a call — the
            // forged call is what produced `bool is not callable (type: true)`.
            ("isEmpty", "dart.is_empty"),
            ("isNotEmpty", "dart.is_not_empty"),
        ],
        _ => &[],
    };
    getters
        .iter()
        .map(|(member, emit)| {
            (
                member.to_string(),
                NamespaceNode::Property {
                    get: Some(Box::new(NamespaceNode::CommonEmit(emit.to_string()))),
                    set: None,
                },
            )
        })
        .collect()
}

/// A type node is only reachable for a receiver whose STATIC TYPE NAMES it.
///
/// That is what splits the two carriers, and it is not a limitation to route
/// around. `var buf = StringBuffer()` infers through `ExprKind::New` to the
/// class name, which normalises to `stringbuffer` and finds this node — so a
/// NAMED type belongs here. `var nums = [9, 9, 9]` infers to the array-shape
/// spelling `int()`, which names no type and never will; a built-in receiver
/// is CLASSIFIED (`builtin_type_of` → `BuiltinType`), not named, and its
/// declared carrier is `[builtin_slots.<type>]` in the profile.
///
/// Registering `string`/`list`/`map`/`set` as type nodes here was measured
/// inert for exactly that reason, and it was a second mechanism for a job the
/// profile already answers. Add a type node when the type has a NAME the
/// inference can produce; declare a builtin slot otherwise.
fn core_types() -> Subtree {
    let mut core = Subtree::new();
    for (name, _) in crate::core_classes::CORE_CLASSES {
        let member_returns = match *name {
            "StringBuffer" => [
                // Dart's own spelling — these are the keys the tree is asked
                // for, and Dart does not fold.
                ("toString", "String"),
                ("length", "int"),
                ("isEmpty", "bool"),
                ("isNotEmpty", "bool"),
            ]
            .as_slice(),
            _ => &[],
        };
        core.insert(
            // The DECLARED spelling — `StringBuffer`, not `stringbuffer`.
            name.to_string(),
            NamespaceNode::Type {
                ctor: None,
                ctor_call: None,
                statics: Subtree::new(),
                methods: core_properties(name),
                member_returns: member_returns
                    .iter()
                    .map(|(k, v)| (k.to_string(), v.to_string()))
                    .collect(),
            },
        );
    }
    core
}

/// The `dart:` library each type belongs to — Dart's own structure, stated
/// once.
///
/// A dotted `[builtins]` key (`"Uri.parse"`, `"Map.from"`, `"int.parse"`) is a
/// TYPE and its STATIC member written as one string. That is a namespace
/// spelled as a flat name-table key: nothing can walk it, `Uri` has no
/// existence apart from the 4 keys that mention it, and a language asking for
/// `dart.core.Uri.parse` finds nothing. Splitting the key at the dot and
/// registering the member under its owner's `statics` is what turns the table
/// back into the namespace it was always describing.
///
/// `math` is deliberately absent: `dart:math` is a LIBRARY, not a type, and it
/// already registers as an alias onto `ecma.math`.
fn library_of(owner: &str) -> &'static str {
    match owner {
        "Future" | "Stream" | "Promise" => "async",
        "Queue" | "LinkedHashMap" | "LinkedHashSet" => "collection",
        "Platform" => "io",
        _ => "core",
    }
}

/// Insert `member` into `owner`'s `statics`, creating the type node if the
/// class-backed pass did not already declare it.
fn insert_static(libraries: &mut Subtree, owner: &str, member: &str, node: NamespaceNode) {
    let library = libraries
        .entry(library_of(owner).to_string())
        .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
    let NamespaceNode::Namespace(types) = library else {
        return;
    };
    let entry = types
        // The DECLARED spelling of the owning type.
        .entry(owner.to_string())
        .or_insert_with(|| NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: Subtree::new(),
            member_returns: BTreeMap::new(),
        });
    let NamespaceNode::Type { statics, .. } = entry else {
        return;
    };
    statics.entry(member.to_string()).or_insert(node);
}

fn insert_adapter_type(libraries: &mut Subtree, adapter: AdapterType) {
    let library = libraries
        .entry(adapter.library.to_string())
        .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
    let NamespaceNode::Namespace(types) = library else {
        return;
    };
    let entry = types
        .entry(adapter.name.to_string())
        .or_insert_with(|| NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: Subtree::new(),
            member_returns: BTreeMap::new(),
        });
    let NamespaceNode::Type { ctor_call, .. } = entry else {
        return;
    };
    if ctor_call.is_none() {
        *ctor_call = Some(Box::new(adapter_ctor_node(adapter.ctor)));
    }
}

/// Register the Dart surface under the `dart` root. Idempotent; first
/// call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let Ok(profile) = parse_profile(super::profile_source()) else {
            return;
        };
        let mut root = Subtree::new();
        for (name, def) in &profile.builtins {
            // The builtin's own spelling — `double.parse`, not `double.parse`
            // folded. Dart is case-sensitive; the tree now keeps the case and
            // folds only at lookup, and only for a caller that asked to.
            let key = name.to_string();
            match &def.emit {
                BuiltinEmit::Common(op) => {
                    if let Some(path) = op.strip_prefix("dart.") {
                        insert_path(&mut root, path, NamespaceNode::CommonEmit(op.clone()));
                    }
                }
                BuiltinEmit::HostCall(module, func) => {
                    insert_path(&mut root, &key, namespaces::host_fn(module, func));
                }
                _ => {}
            }
        }
        for default in &profile.esm_defaults {
            if let EsmDefault::Namespace { alias, module } = default {
                insert_path(
                    &mut root,
                    alias,
                    NamespaceNode::Alias(module_tree_path(module)),
                );
            }
        }
        // ── The `dart:` libraries ────────────────────────────────────────
        //
        // Class-backed types first, so a dotted `[builtins]` key adds statics
        // to the type the AST already declares rather than creating a second,
        // empty one under the same name.
        let mut libraries = Subtree::new();
        libraries.insert("core".to_string(), NamespaceNode::Namespace(core_types()));
        for adapter in ADAPTER_TYPES {
            insert_adapter_type(&mut libraries, *adapter);
        }
        for (name, def) in &profile.builtins {
            let Some((owner, member)) = name.split_once('.') else {
                continue;
            };
            // `math.sqrt` and friends belong to the `dart:math` LIBRARY, which
            // already resolves as an alias onto `ecma.math`.
            if owner == "math" {
                continue;
            }
            let node = match &def.emit {
                BuiltinEmit::Common(op) => NamespaceNode::CommonEmit(op.clone()),
                BuiltinEmit::HostCall(module, func) => namespaces::host_fn(module, func),
                _ => continue,
            };
            insert_static(&mut libraries, owner, member, node);
        }
        for (library, types) in libraries {
            root.insert(library, types);
        }
        namespaces::register_namespace_tree("dart", NamespaceNode::Namespace(root));
    });
}
