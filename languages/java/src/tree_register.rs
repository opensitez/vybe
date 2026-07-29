//! `java.*` namespace-tree registration (namespaceplan.md Phase 5 shape).
//!
//! Mirrors the dotnet registrar: the Java language contributes DATA — its
//! profile `[builtins]` table, the same table its emit dispatch executes —
//! to the shared namespace tree. Resolution LOGIC lives only in the
//! common resolver; any language can walk `java.integer.parseint`.
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

use vybe_bytecode::namespaces::{self, NamespaceNode, Subtree};
use vybe_bytecode::profile::{BuiltinEmit, parse_profile};

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

/// One Java stdlib type, as DATA: what it is called, what it IS (its
/// `is`/`isInstance`/`instanceof` ancestry — self first, then supertypes and
/// interfaces), and — when the type has no object identity to carry a stamp —
/// which runtime intrinsic backs it.
pub(crate) struct JavaType {
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
pub(crate) const JAVA_TYPES: &[JavaType] = &[
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
pub(crate) fn array_element_intrinsic(type_name: &str) -> Option<&'static str> {
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
pub(crate) fn intrinsics_answering(queried: &str) -> Vec<&'static str> {
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
        let Ok(profile) = parse_profile(super::profile_source()) else {
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
            let Some((module, func)) = profile.lookup_known_type(ty.name) else {
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

        for (name, def) in &profile.builtins {
            let key = name.to_lowercase();
            // Internal walker-support helpers are not surface.
            if key.starts_with("__") {
                continue;
            }
            match &def.emit {
                BuiltinEmit::Common(op) => {
                    if key.starts_with("java.") || op.starts_with("java.") {
                        insert_path(&mut root, &key, NamespaceNode::CommonEmit(op.clone()));
                    }
                }
                BuiltinEmit::HostCall(module, func) => {
                    insert_path(&mut root, &key, namespaces::host_fn(module, func));
                }
                _ => {}
            }
        }
        namespaces::register_namespace_tree("java", NamespaceNode::Namespace(root));
    });
}
