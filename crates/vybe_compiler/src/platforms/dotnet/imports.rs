//! `.NET` adapter implicit import list.
//!
//! Owns the single piece of data that represents "what `.NET`-shaped namespaces
//! does every adapter-aware compiler implicitly recognise". Language compilers
//! extend this with language-specific additions (e.g. `microsoft.visualbasic`
//! for VB, `system.linq` for C#) before handing it to the resolver.
//!
//! This file is intentionally narrow — namespace-root recognition lives in
//! `namespaces.rs`, and namespace → host mapping lives in `host_map.rs`.

use std::collections::BTreeSet;

/// Return the default set of `.NET`-shaped namespace imports that every
/// adapter-aware compiler should recognise. Returned as a Vec so callers can
/// `.extend()` with extras.
pub fn default_interface_imports() -> Vec<String> {
    super::core::default_interface_imports()
        .into_iter()
        .chain(super::winforms::default_interface_imports())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_interface_imports_merges_core_and_winforms() {
        let imports = default_interface_imports();
        assert!(imports.contains(&"system".to_string()));
        assert!(imports.contains(&"system.windows.forms".to_string()));
        assert!(imports.contains(&"application".to_string()));
    }
}
