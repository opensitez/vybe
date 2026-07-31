//! `java.*` namespace-tree registration — the JDK as a PLATFORM.
//!
//! Mirrors the dotnet registrar: this crate contributes DATA — its own
//! `java.*` profile fragment — to the shared namespace tree. Resolution
//! LOGIC lives only in the common resolver, so ANY language can walk
//! `java.util.objects.equals` by declaring tree data in its profile, exactly
//! as csharp/vb reach `dotnet.*` with zero `System.*` entries of their own.
//!
//! This used to live in `languages/java` and register through the LANGUAGE
//! hook, which made the JDK the property of one frontend.
//!
//! Leaf rules (dotnet template):
//! - Java package-surface common emits register as `CommonEmit` leaves at
//!   the builtin's own (dotted) key path (`java.util.Objects.equals` →
//!   `java.util.objects.equals`), even when the actual common op is a
//!   shared category such as `object.equals`;
//! - Java shorthand common emits (`Integer.parseInt`) register when the
//!   target is Java-owned (`common:java.<op>`);
//! - host-backed builtins register as `Fn` leaves at their key path;
//! - opcode/intrinsic/print builtins have no process-global target to
//!   point at — skipped.

use std::sync::Once;

use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};
use vybe_runtime::profile::{BuiltinEmit, parse_profile};

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
        cursor = match entry {
            NamespaceNode::Namespace(children) => children,
            // A TYPE used as a namespace — `java.util.BitSet.valueOf` descends
            // THROUGH the `BitSet` type node to its statics. `resolve_segments`
            // already walks it that way; the insert was missing the symmetric
            // case, so every static under a registered type was dropped
            // SILENTLY (the arm below returns without a word). Types register
            // first, so this hit `BitSet.valueOf` the moment `BitSet` gained a
            // `JAVA_TYPES` row.
            NamespaceNode::Type { statics, .. } => statics,
            _ => return, // genuine leaf/namespace collision: first wins
        };
    }
    cursor.entry(leaf.to_string()).or_insert(node);
}


/// Make the node at `path` a `Type`, preserving whatever is already there.
///
/// `type_scopes` resolution (`find_type_node`) only descends through
/// `Namespace` nodes and matches `Type` nodes, so a class whose members live
/// under a plain `Namespace` is invisible to it. Promoting keeps the children
/// as the type's `statics`, which is exactly where a static member belongs.
fn ensure_type_node(root: &mut Subtree, path: &str) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let leaf = segments.pop().expect("non-empty path");
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        cursor = match entry {
            NamespaceNode::Namespace(children) => children,
            NamespaceNode::Type { statics, .. } => statics,
            _ => return,
        };
    }
    let Some(existing) = cursor.get_mut(leaf) else {
        return;
    };
    if let NamespaceNode::Namespace(children) = existing {
        let statics = std::mem::take(children);
        *existing = NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics,
            methods: Subtree::new(),
            member_returns: Default::default(),
        };
    }
}

/// One Java stdlib type, as DATA: what it is called, what it IS (its
/// `is`/`isInstance`/`instanceof` ancestry — self first, then supertypes and
/// interfaces), and — when the type has no object identity to carry a stamp —
/// which runtime intrinsic backs it.
pub struct JavaType {
    pub name: &'static str,
    /// The package the type lives in. Java packages ARE its namespaces, so
    /// this is where the type registers in the tree — `java.util.arraylist`.
    /// One registration then answers both `ArrayList` (matched on the leaf)
    /// and `java.util.ArrayList` (matched on the full path).
    pub package: &'static str,
    pub ancestry: &'static [&'static str],
    /// The SHARED canonical type name that `ExprKind::IsType` already tests
    /// for (`"string"`, `"number"`, `"boolean"`, `"list"`, `"object"`).
    ///
    /// `Some` means the type's runtime representation is an intrinsic, so it
    /// cannot carry a `__type`/`__types` stamp and identity has to be answered
    /// by testing the value's runtime kind. `None` means the value is an
    /// object and answers from the stamped ancestry chain like any user class.
    ///
    /// This is the ONLY Java-specific part of the identity story, and it is a
    /// name, not an emit: the test itself is the shared one.
    pub intrinsic: Option<&'static str>,
}

const fn t(
    name: &'static str,
    package: &'static str,
    ancestry: &'static [&'static str],
    intrinsic: Option<&'static str>,
) -> JavaType {
    JavaType {
        name,
        package,
        ancestry,
        intrinsic,
    }
}

/// The Java types whose IDENTITY the runtime must answer for.
///
/// These are DATA, and they are Java knowledge, so they live in the Java crate.
/// What they are NOT is a second resolution path: registering them as `Type`
/// nodes puts them on the same common-resolver `Ctor` + `ancestry` machinery
/// that dotnet, flutter and plib use, and the `intrinsic` column feeds the
/// walker's normalisation of `X.class.isInstance(v)` / `v instanceof X` into
/// the shared `ExprKind::IsType` — so no Java-specific identity check emits.
///
/// Before this they were registered as bare `CommonEmit` leaves — callable, but
/// with no identity — so a type used as a VALUE (`X.class`) resolved to nothing
/// and the call trapped with "undefined is not callable". The same shape as a
/// platform class that was never materialised.
pub const JAVA_TYPES: &[JavaType] = &[
    // ── Intrinsic-backed: no object to stamp, so identity is a kind test ──
    t(
        "String",
        "lang",
        &[
            "String",
            "CharSequence",
            "Comparable",
            "Serializable",
            "Object",
        ],
        Some("string"),
    ),
    t(
        "Character",
        "lang",
        &["Character", "Comparable", "Serializable", "Object"],
        Some("string"),
    ),
    t(
        "Class",
        "lang",
        &["Class", "Serializable", "Object"],
        Some("string"),
    ),
    t(
        "Integer",
        "lang",
        &["Integer", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Long",
        "lang",
        &["Long", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Short",
        "lang",
        &["Short", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Byte",
        "lang",
        &["Byte", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Double",
        "lang",
        &["Double", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Float",
        "lang",
        &["Float", "Number", "Comparable", "Serializable", "Object"],
        Some("number"),
    ),
    t(
        "Boolean",
        "lang",
        &["Boolean", "Comparable", "Serializable", "Object"],
        Some("boolean"),
    ),
    // The root: its own ancestry is just itself, so it contributes `"object"`
    // to `Object.class.isInstance(…)` without leaking that kind into every
    // interface query. Java's `new Object()` is a bare map at runtime.
    t("Object", "lang", &["Object"], Some("object")),
    // ── Object-backed: identity comes from the stamped ancestry chain ──
    t(
        "ArrayList",
        "util",
        &[
            "ArrayList",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "LinkedList",
        "util",
        &[
            "LinkedList",
            "Deque",
            "Queue",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "ArrayDeque",
        "util",
        &[
            "ArrayDeque",
            "Deque",
            "Queue",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "Vector",
        "util",
        &[
            "Vector",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "Stack",
        "util",
        &[
            "Stack",
            "Vector",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "HashSet",
        "util",
        &[
            "HashSet",
            "Set",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "LinkedHashSet",
        "util",
        &[
            "LinkedHashSet",
            "HashSet",
            "Set",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "TreeSet",
        "util",
        &[
            "TreeSet",
            "NavigableSet",
            "SortedSet",
            "Set",
            "Collection",
            "Iterable",
            "Object",
        ],
        None,
    ),
    t(
        "PriorityQueue",
        "util",
        &["PriorityQueue", "Queue", "Collection", "Iterable", "Object"],
        None,
    ),
    t(
        "HashMap",
        "util",
        &["HashMap", "Map", "Cloneable", "Object"],
        None,
    ),
    t(
        "LinkedHashMap",
        "util",
        &["LinkedHashMap", "HashMap", "Map", "Cloneable", "Object"],
        None,
    ),
    t(
        "TreeMap",
        "util",
        &["TreeMap", "NavigableMap", "SortedMap", "Map", "Object"],
        None,
    ),
    t(
        "StringBuilder",
        "lang",
        &["StringBuilder", "CharSequence", "Appendable", "Object"],
        None,
    ),
    // Types whose `[known_types]` entry declares a constructor but which had no
    // JAVA_TYPES row, so they never registered as tree `Type` nodes. While the
    // declarations lived in the Java profile that was invisible —
    // `lookup_known_type` answered from the language's own table. Once the
    // qualified names moved here, the TREE is the only thing that can answer,
    // so the row is what makes `new java.math.BigInteger(...)` resolve at all.
    t(
        "BigInteger",
        "math",
        &["BigInteger", "Number", "Comparable", "Serializable", "Object"],
        None,
    ),
    t("BitSet", "util", &["BitSet", "Cloneable", "Serializable", "Object"], None),
    t("Random", "util", &["Random", "Serializable", "Object"], None),
    t(
        "SplittableRandom",
        "util",
        &["SplittableRandom", "Object"],
        None,
    ),
    t(
        "Hashtable",
        "util",
        &["Hashtable", "Map", "Cloneable", "Serializable", "Object"],
        None,
    ),
    t(
        "IdentityHashMap",
        "util",
        &["IdentityHashMap", "Map", "Cloneable", "Serializable", "Object"],
        None,
    ),
    t(
        "Properties",
        "util",
        &["Properties", "Hashtable", "Map", "Cloneable", "Object"],
        None,
    ),
    t(
        "WeakHashMap",
        "util",
        &["WeakHashMap", "Map", "Object"],
        None,
    ),
    t(
        "CopyOnWriteArrayList",
        "util.concurrent",
        &[
            "CopyOnWriteArrayList",
            "List",
            "Collection",
            "Iterable",
            "Cloneable",
            "Object",
        ],
        None,
    ),
    t(
        "LinkedBlockingQueue",
        "util.concurrent",
        &[
            "LinkedBlockingQueue",
            "BlockingQueue",
            "Queue",
            "Collection",
            "Iterable",
            "Object",
        ],
        None,
    ),
    t(
        "Semaphore",
        "util.concurrent",
        &["Semaphore", "Serializable", "Object"],
        None,
    ),
    // The one Throwable the profile constructs as a plain map rather than
    // through `ecma:error`. A user class that EXTENDS a built-in exception
    // already gets its chain stamped by the walker; this is the direct
    // `new OutOfMemoryError()` case, which had no identity at all.
    t(
        "OutOfMemoryError",
        "lang",
        &[
            "OutOfMemoryError",
            "VirtualMachineError",
            "Error",
            "Throwable",
            "Object",
        ],
        None,
    ),
];

/// A Java array type and the runtime kind of its ELEMENTS.
///
/// Every Java array is a JS array at runtime and the element type is not
/// represented anywhere on the value, so `String[]` and `int[]` are the same
/// object. The only thing that distinguishes them is what is IN them — hence
/// the element kind, which the caller probes on the first element.
///
/// An array type not listed here (`Object[]`, a generic `T[]`) answers on
/// array-ness alone, which is the most that can be said about it.
pub fn array_element_intrinsic(type_name: &str) -> Option<&'static str> {
    let element = type_name.strip_suffix("[]")?.trim();
    Some(match element {
        "int" | "long" | "short" | "byte" | "double" | "float" => "number",
        "char" | "String" => "string",
        "boolean" => "boolean",
        _ => return None,
    })
}

/// Every distinct runtime intrinsic that could answer `isInstance` for
/// `queried` — the intrinsics of all registered types that HAVE `queried` in
/// their ancestry. `Comparable` yields `["string", "number", "boolean"]`;
/// `String` yields `["string"]`; a user class yields nothing.
///
/// Empty means "no intrinsic can answer this" — the caller falls back to the
/// stamped-ancestry path, which is also what must be OR-ed in alongside a
/// non-empty result so object-backed instances still answer.
pub fn intrinsics_answering(queried: &str) -> Vec<&'static str> {
    let mut out: Vec<&'static str> = Vec::new();
    for ty in JAVA_TYPES {
        let Some(intrinsic) = ty.intrinsic else {
            continue;
        };
        if ty.ancestry.iter().any(|a| *a == queried) && !out.contains(&intrinsic) {
            out.push(intrinsic);
        }
    }
    out
}

/// Register the Java stdlib surface under the `java` root. Idempotent;
/// first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let Ok(profile) = parse_profile(crate::profile_source()) else {
            return;
        };
        let mut root = Subtree::new();

        // TYPES first, at their package path (`java.util.arraylist`), so one
        // registration answers both the bare name and the fully-qualified
        // chain. Their constructor is the one the profile already declares in
        // `[known_types]` — the tree does not invent a second one, it wraps
        // the same target in a `Type` node so construction also STAMPS the
        // declared ancestry and `isInstance` can answer from `__types`.
        //
        // Only object-backed types register. An intrinsic-backed type
        // (`String`, `Integer`) has no object to stamp, so a `Type` node would
        // promise an identity it cannot carry; those answer through the
        // `intrinsic` column instead.
        for ty in JAVA_TYPES.iter().filter(|ty| ty.intrinsic.is_none()) {
            // The QUALIFIED key: this fragment owns `java.util.ArrayList`, not
            // the bare `ArrayList` spelling — that is Java's own idiom and
            // stays in the Java profile. A platform declares package paths.
            let qualified = format!("java.{}.{}", ty.package, ty.name);
            let Some((module, func)) = profile.lookup_known_type(&qualified) else {
                continue;
            };
            let ctor_call = if module == "common" {
                NamespaceNode::CommonEmit(func.to_string())
            } else {
                namespaces::host_fn(module, func)
            };
            let path = format!("{}.{}", ty.package, ty.name.to_lowercase());
            insert_path(
                &mut root,
                &path,
                NamespaceNode::Type {
                    ctor: Some(namespaces::CtorSpec {
                        params: Vec::new(),
                        fields: Vec::new(),
                        ancestry: ty.ancestry.iter().map(|s| (*s).to_string()).collect(),
                        control_fn: None,
                        field_gui: Vec::new(),
                        value_equality: false,
                    }),
                    ctor_call: Some(Box::new(ctor_call)),
                    statics: Subtree::new(),
                    methods: Subtree::new(),
                    member_returns: Default::default(),
                },
            );
        }

        // Package-surface builtins register as STATICS on a `Type` node, the
        // shape dotnet/flutter/plib use — `java.util.Objects.equals` becomes
        // member `equals` on type `java.util.objects`. That is what makes the
        // JDK reachable from ANY language with no declarations of its own:
        // `lookup_type_static_member` / `lookup_type_ctor_target` lowercase
        // BOTH the class name and the member, so `type_scopes = ["java"]`
        // resolves `Objects.equals`, `objects.equals` and every other casing
        // through the common resolver.
        //
        // Registering them as flat leaves at a dotted path — which is what this
        // did — left them reachable only by the plain tree walk, which matches
        // the caller's exact spelling against lowercase keys. Java is
        // case-sensitive, so its own surface was unreachable, and every
        // consumer had to compensate with per-language mount/ambient data.
        //
        // Collected first, then inserted, so a type's statics are built in one
        // go and never depend on `[builtins]` iteration order.
        let mut statics_by_type: std::collections::BTreeMap<String, Subtree> =
            std::collections::BTreeMap::new();

        for (name, def) in &profile.builtins {
            let key = name.to_lowercase();
            // Internal walker-support helpers are not surface.
            if key.starts_with("__") {
                continue;
            }
            // `root` IS the `java` node — it is handed to
            // `register_namespace_tree("java", …)` below. So a key that
            // already carries the package prefix must drop it, or the leaf
            // lands at `java.java.time.instant.parse` and NO resolver walk
            // can reach it. Every `java.*`-keyed builtin was registered that
            // way, which is why the tree's whole package surface was dead
            // while the type nodes (keyed `util.arraylist`) resolved fine.
            let path = key.strip_prefix("java.").unwrap_or(key.as_str());
            let leaf = match &def.emit {
                BuiltinEmit::Common(op) => {
                    if key.starts_with("java.") || op.starts_with("java.") {
                        NamespaceNode::CommonEmit(op.clone())
                    } else {
                        continue;
                    }
                }
                BuiltinEmit::HostCall(module, func) => namespaces::host_fn(module, func),
                _ => continue,
            };
            match path.rsplit_once('.') {
                Some((type_path, member)) => {
                    statics_by_type
                        .entry(type_path.to_string())
                        .or_default()
                        .entry(member.to_string())
                        .or_insert(leaf);
                }
                // A single-segment builtin is package-level, not a member.
                None => insert_path(&mut root, path, leaf),
            }
        }

        for (type_path, statics) in statics_by_type {
            // `or_insert` never clobbers a type already registered from
            // `JAVA_TYPES` above — that one carries the ctor and ancestry, and
            // `insert_path` merges these statics into it.
            for (member, leaf) in statics {
                insert_path(&mut root, &format!("{type_path}.{member}"), leaf);
            }
            ensure_type_node(&mut root, &type_path);
        }
        namespaces::register_namespace_tree("java", NamespaceNode::Namespace(root));
    });
}

#[cfg(test)]
mod tests {
    #[test]
    fn profile_fragment_parses_and_declares_the_surface() {
        let profile = vybe_runtime::profile::parse_profile(crate::profile_source())
            .expect("java.* profile fragment must parse");
        let builtins = profile
            .builtins
            .keys()
            .filter(|k| k.starts_with("java."))
            .count();
        let known = ["java.util.ArrayList", "java.util.HashMap", "java.math.BigInteger"];
        for k in known {
            assert!(
                profile.lookup_known_type(k).is_some(),
                "known_type {k} missing from the platform fragment"
            );
        }
        // 268 distinct JDK statics. Java declared 40 of them TWICE — once
        // qualified (`java.time.Instant.parse`) and once bare
        // (`Instant.parse`) — which is the duplication this crate removes.
        assert_eq!(builtins, 242, "java.* builtins in the platform fragment");
    }
}
