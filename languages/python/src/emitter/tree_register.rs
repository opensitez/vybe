//! Python namespace-tree registration.
//!
//! Mirrors `languages/php/src/tree_register.rs`: the LANGUAGE contributes DATA
//! — its own profile tables, the same ones its emit dispatch executes — to the
//! shared tree in `vybe_runtime::namespaces`. Resolution logic lives only in
//! the common resolver; nothing in the VM or the common compiler changes.
//!
//! The one difference from PHP is what a DOTTED profile key means. PHP skips
//! them (`$obj->method` is receiver dispatch), but in Python `"os.stat"`,
//! `"os.path.join"` and `"math.pi"` ARE module members, so each dotted key
//! nests into real subtrees — `os` → `path` → `join`. That is what makes
//! `from os import stat` / `from math import pi` resolve generically, instead
//! of needing a hand-written `[[esm_default]] kind = "module-export"` row per
//! name (the json rows are exactly that workaround).
//!
//! Leaf kinds come from the profile entry:
//! - `emit = "common:python.<fn>"` → `CommonEmit`
//! - `emit = "host:<module>:<fn>"` → `Fn`
//! - `[namespace_constants]` → `Const`
//! - opcode/intrinsic/print builtins have no process-global target — skipped.
//!
//! Python is CASE-SENSITIVE, so keys keep their exact source casing
//! (`OrderedDict`, `NamedTemporaryFile`); the lowercase-canonical rule applies
//! only to case-insensitive languages, so the `Subtree` is built directly
//! rather than through the lowercase-asserting `namespace()` helper.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_runtime::Value;
use vybe_runtime::namespaces::{self, CtorSpec, NamespaceNode, Subtree};
use vybe_runtime::profile::{BuiltinEmit, ConstantValue, parse_profile};

/// Insert `leaf` at a dotted path (`["path", "join"]` under root `os`),
/// creating intermediate namespaces. An existing entry wins — first
/// registration is authoritative, as in the other registrars.
fn insert_path(root: &mut Subtree, path: &[&str], leaf: NamespaceNode) {
    let Some((last, parents)) = path.split_last() else {
        return;
    };
    let mut current = root;
    for seg in parents {
        let entry = current
            .entry((*seg).to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        match entry {
            NamespaceNode::Namespace(children) => current = children,
            // A leaf already occupies this segment — it cannot also be a
            // namespace, so leave the working entry alone.
            _ => return,
        }
    }
    current.entry((*last).to_string()).or_insert(leaf);
}

/// Register the Python surface. Idempotent; first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        register_collections();
        register_calendar();
        register_from_profile();
    });
}

/// Everything derivable from the profile — no hand-maintained module list.
fn register_from_profile() {
    let Ok(profile) = parse_profile(crate::profile_source()) else {
        return;
    };

    let mut roots: BTreeMap<String, Subtree> = BTreeMap::new();
    let mut add = |key: &str, leaf: NamespaceNode| {
        let segments: Vec<&str> = key.split('.').collect();
        // A bare builtin (`len`, `print`) is not module surface.
        if segments.len() < 2 {
            return;
        }
        let root = roots.entry(segments[0].to_string()).or_default();
        insert_path(root, &segments[1..], leaf);
    };

    for (name, def) in &profile.builtins {
        match &def.emit {
            BuiltinEmit::Common(op) => add(name, NamespaceNode::CommonEmit(op.clone())),
            BuiltinEmit::HostCall(module, func) => add(name, namespaces::host_fn(module, func)),
            _ => {}
        }
    }

    // `[namespace_constants]` are VALUES (`math.pi`, `math.inf`) — precisely
    // the members a named import binds. Without them `from math import pi` has
    // nothing to resolve to and lands as `nan`.
    for (name, value) in &profile.namespace_constants {
        let node = match value {
            ConstantValue::Bool(b) => NamespaceNode::Const(Value::Bool(*b)),
            ConstantValue::Float(f) => NamespaceNode::Const(Value::F64(*f)),
            // `inf`/`nan` are spelled as strings in the profile but ARE floats.
            ConstantValue::Str(s) => match s.as_str() {
                "Infinity" => NamespaceNode::Const(Value::F64(f64::INFINITY)),
                "-Infinity" => NamespaceNode::Const(Value::F64(f64::NEG_INFINITY)),
                "NaN" => NamespaceNode::Const(Value::F64(f64::NAN)),
                _ => NamespaceNode::Const(Value::String(std::sync::Arc::from(s.as_str()))),
            },
        };
        add(name, node);
    }

    // ECMA-shared math names. The host exports them under `ecma:math`, which
    // `mount_host_exports` mounts LOWERCASED as `ecma.math.<name>` — so an
    // `Alias` leaf reaches them with no per-name profile row. namespaceplan.md
    // §"Source-name ≠ canonical-name": `python.json.dumps =
    // Alias(ecma.json.stringify)` is exactly this shape. These five are the
    // only `FLOAT_MATH_FNS` members with no profile entry.
    for name in ["hypot", "gamma", "lgamma", "erf", "erfc"] {
        // `Path` is a dotted string.
        let target = match name {
            // tgamma/lgamma live in the libc platform tree, not ecma.
            "gamma" => "libc.math.tgamma".to_string(),
            "lgamma" => "libc.math.lgamma".to_string(),
            _ => format!("ecma.math.{name}"),
        };
        let root = roots.entry("math".to_string()).or_default();
        insert_path(root, &[name], NamespaceNode::Alias(target));
    }

    // Register each module as its own root (`math.pi`, `os.stat`) AND under a
    // `python.*` package root, so ANY language can walk `python.math.pi` the
    // same way it can walk `php.str_replace`. Same data, two mount points.
    let mut python_root = Subtree::new();
    for (root, tree) in &roots {
        python_root.insert(root.clone(), NamespaceNode::Namespace(tree.clone()));
    }
    namespaces::register_namespace_tree("python", NamespaceNode::Namespace(python_root));

    for (root, tree) in roots {
        namespaces::register_namespace_tree(&root, NamespaceNode::Namespace(tree));
    }
}

/// `collections` constructors. Hand-written because these are not profile
/// builtins: `deque`/`OrderedDict` ARE plain ecma constructors, and
/// `Counter`/`defaultdict` need custom construction.
///
/// Instance methods (`rotate`, `move_to_end`, `most_common`, …) are NEVER
/// resolved through the tree — member dispatch is receiver-based.
fn register_collections() {
    let mut root = Subtree::new();
    root.insert(
        "deque".to_string(),
        namespaces::host_fn("ecma:array", "from"),
    );
    root.insert(
        "OrderedDict".to_string(),
        namespaces::host_fn("ecma:object", "create"),
    );
    root.insert(
        "Counter".to_string(),
        NamespaceNode::CommonEmit("python.counter_new".to_string()),
    );
    root.insert(
        "defaultdict".to_string(),
        NamespaceNode::CommonEmit("python.defaultdict_new".to_string()),
    );
    namespaces::register_namespace_tree("collections", NamespaceNode::Namespace(root));
}

fn calendar_type(name: &str, ctor: &str, formatmonth: Option<&str>) -> NamespaceNode {
    let mut methods = BTreeMap::from([
        (
            "itermonthdays".to_string(),
            NamespaceNode::CommonEmit("python.calendar_itermonthdays".to_string()),
        ),
        (
            "itermonthdates".to_string(),
            NamespaceNode::CommonEmit("python.calendar_itermonthdays".to_string()),
        ),
        (
            "itermonthdays2".to_string(),
            NamespaceNode::CommonEmit("python.calendar_itermonthdays2".to_string()),
        ),
        (
            "monthdayscalendar".to_string(),
            NamespaceNode::CommonEmit("python.calendar_monthcalendar".to_string()),
        ),
        (
            "yeardayscalendar".to_string(),
            NamespaceNode::CommonEmit("python.calendar_yeardayscalendar".to_string()),
        ),
    ]);
    if let Some(emit) = formatmonth {
        methods.insert(
            "formatmonth".to_string(),
            NamespaceNode::CommonEmit(emit.to_string()),
        );
        methods.insert(
            "prmonth".to_string(),
            NamespaceNode::CommonEmit(emit.to_string()),
        );
    }
    NamespaceNode::Type {
        ctor: Some(CtorSpec {
            params: vec!["firstweekday".to_string()],
            fields: vec!["firstweekday".to_string()],
            ancestry: vec![name.to_string()],
            ..Default::default()
        }),
        ctor_call: Some(Box::new(NamespaceNode::CommonEmit(ctor.to_string()))),
        statics: Subtree::new(),
        methods,
        member_returns: BTreeMap::new(),
    }
}

fn register_calendar() {
    let mut root = Subtree::new();
    root.insert(
        "Calendar".to_string(),
        calendar_type("Calendar", "python.calendar_new", None),
    );
    root.insert(
        "TextCalendar".to_string(),
        calendar_type(
            "TextCalendar",
            "python.calendar_text_new",
            Some("python.calendar_text_formatmonth"),
        ),
    );
    root.insert(
        "HTMLCalendar".to_string(),
        calendar_type(
            "HTMLCalendar",
            "python.calendar_html_new",
            Some("python.calendar_html_formatmonth"),
        ),
    );
    namespaces::register_namespace_tree("calendar", NamespaceNode::Namespace(root));
}
