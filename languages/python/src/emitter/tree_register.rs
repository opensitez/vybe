//! `collections.*` namespace-tree registration — the Python stdlib
//! `collections` module contributed to the SHARED common resolver, exactly like
//! `dotnet.*` (C#) and `php.*` (PHP). Resolution logic lives only in the common
//! resolver; `from collections import deque` / `collections.deque(...)` walk the
//! tree via `resolve_path(["collections", "deque"])`.
//!
//! Python is CASE-SENSITIVE, so keys keep their exact source casing
//! (`OrderedDict`, `Counter`) — the lowercase-canonical rule applies only to
//! case-insensitive languages, so we build the `Subtree` directly rather than
//! via the lowercase-asserting `namespace()` helper.
//!
//! Leaf kinds:
//! - plain host-backed ctors register as `Fn` leaves (`deque`/`OrderedDict` are
//!   just an ecma array / ecma object);
//! - ctors needing custom construction logic register as `CommonEmit` leaves
//!   dispatched through the Python emit dispatcher (`Counter`, `defaultdict`).
//!
//! Instance methods (`rotate`, `move_to_end`, `most_common`, …) are NEVER
//! resolved through the tree — member dispatch is receiver-based (the profile
//! method table + emit dispatch). Only the constructors live here.

use std::sync::Once;

use vybe_emitter::namespaces::{self, NamespaceNode, Subtree};

/// Register the `collections` module surface under the `collections` root.
/// Idempotent; first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut root = Subtree::new();
        // `deque(iterable)` IS an ecma array (same as `list`); its methods
        // (append/pop/appendleft/popleft/rotate/extendleft) are array emits.
        root.insert("deque".to_string(), namespaces::host_fn("ecma:array", "from"));
        // `OrderedDict(...)` IS an insertion-ordered dict — an ecma object.
        root.insert(
            "OrderedDict".to_string(),
            namespaces::host_fn("ecma:object", "create"),
        );
        // Counting / default-factory maps need custom construction.
        root.insert(
            "Counter".to_string(),
            NamespaceNode::CommonEmit("python.counter_new".to_string()),
        );
        root.insert(
            "defaultdict".to_string(),
            NamespaceNode::CommonEmit("python.defaultdict_new".to_string()),
        );
        namespaces::register_namespace_tree("collections", NamespaceNode::Namespace(root));
    });
}
