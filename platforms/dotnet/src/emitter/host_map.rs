//! .NET → Vybe host mapping tables.
//!
//! The `.NET` BCL surface (`System.Console.WriteLine`, `Math.Sqrt`, etc.) is
//! exposed to user code via Vybe host functions (`wasi:cli::log`,
//! `ecma:math::sqrt`, etc.). This file owns BOTH translation tables that
//! make that work:
//!
//! `static_method_mappings` is what survives of that: the static `.NET` class
//! members routed to a host function, collected from the core and WinForms
//! tables. Member resolution itself goes through the namespace TREE
//! (`tree_register`), not a dotted-name cascade.

use std::sync::LazyLock;

/// Static `.NET` class member routed to a host function.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DotnetStaticMethodMapping {
    pub interface: &'static str,
    pub type_name: &'static str,
    pub method_name: &'static str,
    pub host_module: &'static str,
    pub host_fn: &'static str,
    pub arity: u8,
}

static STATIC_METHOD_MAPPINGS: LazyLock<Vec<DotnetStaticMethodMapping>> = LazyLock::new(|| {
    super::core::static_method_mappings()
        .iter()
        .chain(super::winforms::static_method_mappings())
        .copied()
        .collect()
});

pub fn static_method_mappings() -> &'static [DotnetStaticMethodMapping] {
    STATIC_METHOD_MAPPINGS.as_slice()
}


#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_static_method_mappings_merge_core_and_winforms() {
        assert!(
            static_method_mappings()
                .iter()
                .any(|mapping| mapping.type_name == "Convert")
        );
        assert!(
            static_method_mappings()
                .iter()
                .all(|mapping| mapping.type_name != "Application")
        );
    }
}
