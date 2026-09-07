//! Go namespace-tree registration.
//!
//! Go package names are namespace surface, not prelude globals. This registrar
//! contributes the `go.*` tree from Go-owned adapters so imports and
//! package-qualified names resolve through the shared namespace resolver.

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let Some(leaf) = segments.pop() else {
        return;
    };
    if leaf.is_empty() {
        return;
    }

    let mut cursor = root;
    for seg in segments {
        if seg.is_empty() {
            return;
        }
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.entry(leaf.to_string()).or_insert(node);
}

fn ensure_namespace(root: &mut Subtree, path: &str) {
    let mut cursor = root;
    for seg in path.split('.') {
        if seg.is_empty() {
            return;
        }
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
}

fn add_package_aliases(root: &mut Subtree) {
    for package in [
        "fmt",
        "strings",
        "strconv",
        "math",
        "bits",
        "sort",
        "time",
        "os",
        "rand",
        "path",
        "reflect",
        "net.url",
        "netip",
        "encoding.base64",
        "encoding.binary",
        "encoding.hex",
        "encoding.json",
        "hash.crc32",
        "hash.adler32",
        "hash.fnv",
        "unicode.utf8",
        "unicode.utf16",
        "container.list",
        "container.ring",
        "container.heap",
        "sync",
        "atomic",
        "slices",
        "maps",
        "iter",
        "encoding.xml",
        "encoding.gob",
        "log",
        "log.slog",
        "flag",
        "errors",
    ] {
        ensure_namespace(root, package);
    }

    let aliases = [
        ("filepath", "path"),
        ("url", "net.url"),
        ("netip", "netip"),
        ("base64", "encoding.base64"),
        ("binary", "encoding.binary"),
        ("hex", "encoding.hex"),
        ("json", "encoding.json"),
        ("xml", "encoding.xml"),
        ("gob", "encoding.gob"),
        ("crc32", "hash.crc32"),
        ("adler32", "hash.adler32"),
        ("fnv", "hash.fnv"),
        ("utf8", "unicode.utf8"),
        ("utf16", "unicode.utf16"),
        ("list", "container.list"),
        ("ring", "container.ring"),
        ("heap", "container.heap"),
        ("slog", "log.slog"),
    ];

    for (alias, target) in aliases {
        insert_path(root, alias, NamespaceNode::Alias(format!("go.{target}")));
    }
}

pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut root: Subtree = BTreeMap::new();
        add_package_aliases(&mut root);
        crate::adapters::atomic::register_tree(&mut root);
        crate::adapters::bytes_io::register_tree(&mut root);
        crate::adapters::bits::register_tree(&mut root);
        crate::adapters::container::register_tree(&mut root);
        crate::adapters::encoding::register_tree(&mut root);
        crate::adapters::errors::register_tree(&mut root);
        crate::adapters::flags::register_tree(&mut root);
        crate::adapters::formatting::register_tree(&mut root);
        crate::adapters::gob::register_tree(&mut root);
        crate::adapters::hash::register_tree(&mut root);
        crate::adapters::iter::register_tree(&mut root);
        crate::adapters::json::register_tree(&mut root);
        crate::adapters::logging::register_tree(&mut root);
        crate::adapters::math::register_tree(&mut root);
        crate::adapters::netip::register_tree(&mut root);
        crate::adapters::os::register_tree(&mut root);
        crate::adapters::path::register_tree(&mut root);
        crate::adapters::rand::register_tree(&mut root);
        crate::adapters::reflect::register_tree(&mut root);
        crate::adapters::slices_maps::register_tree(&mut root);
        crate::adapters::sort::register_tree(&mut root);
        crate::adapters::strconv::register_tree(&mut root);
        crate::adapters::strings::register_tree(&mut root);
        crate::adapters::sync::register_tree(&mut root);
        crate::adapters::time::register_tree(&mut root);
        crate::adapters::unicode::register_tree(&mut root);
        crate::adapters::url::register_tree(&mut root);
        crate::adapters::xml::register_tree(&mut root);

        namespaces::register_namespace_tree("go", NamespaceNode::Namespace(root));
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_compiler::primitives::namespaces::ResolutionTarget;
    use vybe_runtime::Value;

    #[test]
    fn go_adapter_surface_registers_under_go_root() {
        register_namespace_tree();

        assert!(matches!(
            namespaces::resolve_path(&["go", "fmt", "Println"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.fmt_println"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "math", "Sqrt"], None),
            Some(ResolutionTarget::HostCall { module, func, .. })
                if module == "ecma:math" && func == "sqrt"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "time", "Hour"], None),
            Some(ResolutionTarget::Const(Value::F64(v))) if v == 3600000000000.0
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "os", "Getenv"], None),
            Some(ResolutionTarget::HostCall { module, func, .. })
                if module == "wasi:cli" && func == "getEnv"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "rand", "Intn"], None),
            Some(ResolutionTarget::HostCall { module, func, .. })
                if module == "ecma:math" && func == "random"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "reflect", "TypeOf"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.reflect_typeof"
        ));
        assert!(matches!(
            namespaces::lookup_type_instance_member(
                &["go".to_string()],
                "__goReflectValue",
                "FieldByName",
                None
            ),
            Some(NamespaceNode::CommonEmit(op)) if op == "go.reflect_field_by_name"
        ));
        assert!(matches!(
            namespaces::lookup_type_instance_member(
                &["go".to_string()],
                "__goReflectType",
                "Kind",
                None
            ),
            Some(NamespaceNode::CommonEmit(op)) if op == "go.reflect_kind"
        ));
    }

    #[test]
    fn go_import_local_names_mount_to_package_tree() {
        register_namespace_tree();

        assert!(matches!(
            namespaces::resolve_path(&["go", "bits", "OnesCount"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.bits_ones_count"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "path"], None),
            Some(ResolutionTarget::NamespaceObject(_))
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "filepath"], None),
            Some(ResolutionTarget::NamespaceObject(_))
        ));
    }

    #[test]
    fn go_atomic_types_register_as_tree_types() {
        register_namespace_tree();

        assert!(matches!(
            namespaces::resolve_path(&["go", "atomic", "Int64"], None),
            Some(ResolutionTarget::Ctor { .. })
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "atomic", "Value"], None),
            Some(ResolutionTarget::Ctor { .. })
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "slices", "Contains"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "collections.contains"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "maps", "Copy"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "collections.map_copy"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "maps", "Keys"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.maps_keys"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "maps", "Equal"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.maps_equal"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "slices", "Values"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.slices_values"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "sort", "Search"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.sort_search"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "sort", "SearchInts"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.sort_search_ordered"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "sort", "Reverse"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.sort_reverse"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "sort", "Ints"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "collections.sort"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "time", "Now"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.time_now"
        ));
        assert!(matches!(
            namespaces::lookup_type_instance_member(
                &["go".to_string()],
                "time.Duration",
                "Minutes",
                None
            ),
            Some(NamespaceNode::CommonEmit(op)) if op == "go.dur_minutes"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "filepath", "Clean"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.path_clean"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "errors", "New"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.errors_new"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "strings", "ReplaceAll"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.strings.ReplaceAll"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "strings", "Contains"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "str_contains"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "strings", "Split"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "strings.split"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "container", "list", "New"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.container.list.New"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "netip", "ParseAddr"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.netip.ParseAddr"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "xml", "EscapeText"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.encoding.xml.EscapeText"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "xml", "NewDecoder"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.encoding.xml.NewDecoder"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "xml", "NewEncoder"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.encoding.xml.NewEncoder"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "gob", "NewEncoder"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.encoding.gob.NewEncoder"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "gob", "NewDecoder"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.encoding.gob.NewDecoder"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "gob", "RegisterName"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.encoding.gob.RegisterName"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "json", "RawMessage"], None),
            Some(ResolutionTarget::Ctor { path, .. }) if path == "go.encoding.json.RawMessage"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "slog", "Info"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.slog.Info"
        ));
        assert!(matches!(
            namespaces::resolve_path(&["go", "flag", "NewFlagSet"], None),
            Some(ResolutionTarget::CommonEmit(op)) if op == "go.flag.NewFlagSet"
        ));
    }
}
