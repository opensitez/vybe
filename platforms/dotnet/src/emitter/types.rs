//! .NET BCL type tables, predicates, and name-shape helpers.
//!
//! This is the static reference data the rest of the .NET frontend leans on:
//! - `known_types()` — bare type name → host constructor mapping for `New X()`
//! - `is_noop_method` — WinForms layout/lifecycle methods that compile to null
//! - `is_known_constant` — .NET property-like constants (Math.PI, etc.) that
//!   shouldn't be invoked even when args are empty
//! - `capitalize_control_name` / `capitalize_data_type` — name shape helpers
//!   used by callers that need PascalCase forms

use std::collections::HashMap;
use std::sync::LazyLock;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum KnownTypeTarget {
    Host {
        module: &'static str,
        constructor: &'static str,
    },
    Common {
        emit: &'static str,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct KnownTypeMapping {
    pub name: &'static str,
    pub interface: &'static str,
    pub display_name: &'static str,
    pub target: KnownTypeTarget,
}

static KNOWN_TYPE_MAPPINGS: LazyLock<Vec<KnownTypeMapping>> = LazyLock::new(|| {
    super::core::known_type_mappings()
        .iter()
        .chain(super::winforms::known_type_mappings())
        .copied()
        .collect()
});

/// WinForms layout/lifecycle methods that are always no-ops at runtime.
pub fn is_noop_method(name: &str) -> bool {
    super::winforms::is_noop_method(name)
}

/// .NET property-like constants that should NOT be called even when args are empty.
pub fn is_known_constant(name: &str) -> bool {
    super::core::is_known_constant(name)
}

/// Return the .NET constructor table: bare type name → constructor target.
pub fn known_type_mappings() -> &'static [KnownTypeMapping] {
    KNOWN_TYPE_MAPPINGS.as_slice()
}

pub fn lookup_known_type(name: &str) -> Option<&'static KnownTypeMapping> {
    known_type_mappings()
        .iter()
        .find(|mapping| mapping.name.eq_ignore_ascii_case(name))
}

pub fn known_types() -> HashMap<String, KnownTypeTarget> {
    let mut m = HashMap::new();
    for mapping in known_type_mappings() {
        m.insert(mapping.name.to_string(), mapping.target);
    }
    m
}

// ─── Name shape helpers ──────────────────────────────────────────────────────
//
// The .NET surface — `System.Windows.Forms.Button`, `System.Windows.Forms.TextBox`,
// etc. — is one frontend on top of the canonical GUI vocabulary in
// `compiler_common::gui`. The .NET frontend's job is name resolution: take a
// .NET-shaped identifier and return the canonical control name. The actual
// emit (host fn naming, calling convention) lives in `gui.rs`.

/// Capitalise a lowercase WinForms control name to its proper casing.
/// Returns an empty string if the name is not a known control.
///
/// Thin wrapper over `compiler_common::gui::canonical_control_name`, kept for
/// backward compatibility with .NET-specific call sites. New code should call
/// into `gui.rs` directly. The .NET frontend assumes the source uses .NET
/// PascalCase, but the canonical name returned matches what other frontends
/// (MAUI, Flutter, Tkinter, …) would also produce.
pub fn capitalize_control_name(name: &str) -> String {
    super::winforms::capitalize_control_name(name)
}

/// Data table / DataSet / DataAdapter — these are .NET BCL data types, NOT
/// GUI controls. They live in `known_types` because they're .NET-specific;
/// other framework frontends won't have them. Returns empty for non-data types.
pub fn capitalize_data_type(name: &str) -> String {
    super::core::capitalize_data_type(name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_known_type_mappings_merge_core_and_winforms() {
        assert!(
            known_type_mappings()
                .iter()
                .any(|mapping| mapping.name == "stringbuilder")
        );
        assert!(
            known_type_mappings()
                .iter()
                .any(|mapping| mapping.name == "form")
        );
    }
}
