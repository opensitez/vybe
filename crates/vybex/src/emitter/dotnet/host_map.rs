//! .NET → Vybe host mapping tables.
//!
//! The `.NET` BCL surface (`System.Console.WriteLine`, `Math.Sqrt`, etc.) is
//! exposed to user code via Vybe host functions (`wasi:cli::log`,
//! `vybe:math::sqrt`, etc.). This file owns BOTH translation tables that
//! make that work:
//!
//! 1. **`namespace_to_host_module`** — `system.console` → `wasi:cli`
//! 2. **`map_host_func`** — `(wasi:cli, writeline)` → `log`
//!
//! The two are kept together because the second table looks up the host
//! module from the first table when the resolver expands a dotted name into
//! a host call.
//!
//! GUI-specific mappings (`vybe:gui::new_Button` etc.) delegate to
//! `compiler_common::gui::canonical_control_name` so the canonical naming
//! lives in one place across all framework frontends.

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

impl DotnetStaticMethodMapping {
    fn matches_legacy_func(&self, module: &str, func: &str) -> bool {
        if self.host_module != module {
            return false;
        }

        let bare = self.method_name.to_lowercase();
        let qualified = format!("{}.{}", self.type_name.to_lowercase(), bare);
        func == bare || func == qualified
    }
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

/// Map a .NET namespace prefix (lowercased, dot-separated) to the Vybe host
/// module name. Returns the prefix itself if no explicit mapping exists.
pub fn namespace_to_host_module<'a>(prefix: &'a str) -> &'a str {
    super::core::namespace_to_host_module(prefix)
        .or_else(|| super::winforms::namespace_to_host_module(prefix))
        .unwrap_or(prefix)
}

/// Map a (host_module, dotnet_method_name) pair to the actual host function
/// name registered in the VM. Both inputs should already be lowercased.
pub fn map_host_func(module: &str, func: &str) -> String {
    if let Some(mapping) = static_method_mappings()
        .iter()
        .find(|mapping| mapping.matches_legacy_func(module, func))
    {
        return mapping.host_fn.to_string();
    }

    match (module, func) {
        _ => super::core::map_host_func(module, func)
            .or_else(|| super::winforms::map_host_func(module, func))
            .unwrap_or_else(|| func.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_static_method_mappings_merge_core_and_winforms() {
        assert!(static_method_mappings().iter().any(|mapping| mapping.type_name == "Console"));
        assert!(static_method_mappings().iter().any(|mapping| mapping.type_name == "Application"));
    }
}
