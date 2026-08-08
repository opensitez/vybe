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

use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};
use vybe_runtime::profile::{BuiltinEmit, EsmDefault, parse_profile};

/// Insert `node` at the dotted `path` under `root`, creating interior
/// namespaces as needed. Keys are lowercase-canonical.
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
    module.replace([':', '/'], ".").to_lowercase()
}

/// `dart:core` TYPES, declared in the namespace tree.
///
/// The tree declares the NAME — that is what makes `dart.core.StringBuffer`
/// reachable through the common resolver from any language, the same way
/// `flutter.*` and `dotnet.*` are. It deliberately carries **no `ctor_call`
/// and no `methods`**: the class itself is an ordinary `ClassDecl`
/// (`core_classes.rs`) that normalises and compiles like any user class, so
/// construction and member dispatch are already answered by `compile_class`'s
/// rtt, vtable and prototype. A `ctor_call` here would intercept construction
/// and put back the anonymous `struct.new 0` this migration removes; a
/// `methods` entry would win over the class's own member.
///
/// `member_returns` stays, because a declared return type is knowledge the
/// tree owns and the class body does not restate.
///
/// The name list comes from `core_classes::CORE_CLASSES`, so the tree cannot
/// declare a type the AST does not build, or miss one it does.
fn core_types() -> Subtree {
    let mut core = Subtree::new();
    for (name, _) in crate::core_classes::CORE_CLASSES {
        let member_returns = match *name {
            "StringBuffer" => [
                ("tostring", "String"),
                ("length", "int"),
                ("isempty", "bool"),
                ("isnotempty", "bool"),
            ]
            .as_slice(),
            _ => &[] };
        core.insert(
            name.to_lowercase(),
            NamespaceNode::Type {
                ctor: None,
                ctor_call: None,
                statics: Subtree::new(),
                methods: Subtree::new(),
                member_returns: member_returns
                    .iter()
                    .map(|(k, v)| (k.to_string(), v.to_string()))
                    .collect() },
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
        "Queue" => "collection",
        "Platform" => "io",
        _ => "core" }
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
        .entry(owner.to_lowercase())
        .or_insert_with(|| NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: Subtree::new(),
            member_returns: BTreeMap::new() });
    let NamespaceNode::Type { statics, .. } = entry else {
        return;
    };
    statics.entry(member.to_lowercase()).or_insert(node);
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
            let key = name.to_lowercase();
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
                    &alias.to_lowercase(),
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
                _ => continue };
            insert_static(&mut libraries, owner, member, node);
        }
        for (library, types) in libraries {
            root.insert(library, types);
        }
        namespaces::register_namespace_tree("dart", NamespaceNode::Namespace(root));
    });
}
