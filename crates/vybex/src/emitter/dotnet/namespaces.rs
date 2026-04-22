//! .NET namespace root recognition.
//!
//! "Is `Math` a variable or the start of a namespace/type chain?" The
//! resolver consults `is_namespace_root` to disambiguate. The set is
//! computed once at first access via `LazyLock`.
//!
//! Sibling files:
//! - `imports.rs` — the import LIST (what's auto-imported)
//! - `host_map.rs` — the .NET → host fn translation tables
//! - `resolver.rs` — uses `is_namespace_root` during dotted-name resolution

use std::collections::HashSet;
use std::sync::LazyLock;

/// The static set of namespace roots, computed once.
static NAMESPACE_ROOTS: LazyLock<HashSet<String>> = LazyLock::new(|| namespace_roots());

/// Check if a name is a known .NET namespace root.
pub fn is_namespace_root(name: &str) -> bool {
    NAMESPACE_ROOTS.contains(name)
}

/// Return the set of names that should be treated as namespace/type roots.
/// Public so language-specific extensions can build derived sets if needed.
pub fn namespace_roots() -> HashSet<String> {
    let mut roots = super::core::namespace_roots();
    roots.extend(super::winforms::namespace_roots());
    roots
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_namespace_roots_merge_core_and_winforms() {
        let roots = namespace_roots();
        assert!(roots.contains("console"));
        assert!(roots.contains("application"));
        assert!(roots.contains("formborderstyle"));
    }
}
