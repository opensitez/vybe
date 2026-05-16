use std::collections::HashSet;
use std::sync::LazyLock;

static NAMESPACE_ROOTS: LazyLock<HashSet<String>> = LazyLock::new(namespace_roots);

pub fn is_namespace_root(name: &str) -> bool {
    NAMESPACE_ROOTS.contains(name)
}

pub fn namespace_roots() -> HashSet<String> {
    let mut roots = HashSet::new();
    for name in &[
        "application", "window", "messagebox",
        "borderstyle", "formborderstyle", "contentalignment",
        "autoscalemode", "autosizemode",
        "dialogresult", "messageboxbuttons", "messageboxicon",
        "keys", "dockstyle", "anchorstyles", "formstartposition",
        "formwindowstate",
    ] {
        roots.insert(name.to_string());
    }
    roots
}