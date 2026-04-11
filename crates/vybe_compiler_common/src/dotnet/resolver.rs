//! Dotted-name resolution for .NET frontends.
//!
//! When a compiler encounters a dotted name like `sw.ElapsedMilliseconds` or
//! `System.Threading.Thread.Sleep`, it must decide: is this a namespace chain
//! (static access) or an instance member access (local variable)?
//!
//! Resolution follows .NET semantics:
//! 1. **Locals first** — if the first part is a local variable, the rest is
//!    instance member access (struct_get chain on the object)
//! 2. **Module-level fields** — class/module fields (Me.field)
//! 3. **Fully-qualified namespace** — `System.Threading.Thread.Sleep`
//! 4. **Imports-resolved** — `Thread.Sleep` with `Imports System.Threading`
//! 5. **Known type static member** — `String.Format` as a type static method
//!
//! Compilers call `resolve_dotted_name()` which inspects the parts and returns
//! a `DottedResolution` telling the compiler exactly what bytecode to emit.
//!
//! This file owns the resolution algorithm AND its private helpers
//! (`try_resolve_via_imports*`). The translation tables it consults
//! (`namespace_to_host_module`, `map_host_func`) live in `host_map.rs`,
//! the namespace-root recognition lives in `namespaces.rs`, and the
//! noop-method predicate lives in `types.rs`.

use super::host_map::{map_host_func, namespace_to_host_module};
use super::namespaces::is_namespace_root;
use super::types::is_noop_method;

// ─── Resolution result ───────────────────────────────────────────────────────

/// The result of resolving a dotted name chain.
#[derive(Debug, Clone, PartialEq)]
pub enum DottedResolution {
    /// The first part is a local variable. The compiler should:
    /// - emit local_get for the variable
    /// - emit struct_get for each remaining part (instance member access)
    /// No call is implied; the caller decides whether to call or just access.
    InstanceMember {
        /// The local variable name (lowercased)
        local: String,
        /// The remaining parts after the local (lowercased member chain)
        members: Vec<String>,
    },

    /// Resolved to a host import via interface_imports (compile-time resolved).
    /// The compiler should emit call_import(module, func) directly.
    HostCall {
        module: String,
        func: String,
    },

    /// Resolved to a namespace object chain. The compiler should:
    /// - emit global_get for the root namespace
    /// - emit struct_get for each subsequent part
    /// The final value is a callable (for method calls) or a value (for property access).
    NamespaceAccess {
        /// The parts of the fully-qualified chain (lowercased)
        parts: Vec<String>,
    },

    /// A WinForms/layout no-op method. The compiler should emit null and skip.
    NoOp,

    /// Could not resolve — the compiler should fall back to its own logic.
    Unresolved,
}

/// Context provided by the compiler for name resolution.
/// This abstracts over VB/C# differences (different AST types, different scoping).
pub struct ResolutionContext<'a> {
    /// Check whether a name (lowercased) is a local variable in the current scope.
    pub is_local: &'a dyn Fn(&str) -> bool,
    /// Check whether a name (lowercased) is a field of the current class/module.
    pub is_class_field: &'a dyn Fn(&str) -> bool,
    /// Check whether a name is a user-defined class/module name.
    pub is_user_type: &'a dyn Fn(&str) -> bool,
    /// The active import list (e.g. ["system", "system.threading", ...])
    pub imports: &'a [String],
}

// ─── Main entry point ────────────────────────────────────────────────────────

/// Resolve a dotted name chain following .NET resolution order.
///
/// `parts` is the member-access chain (e.g. ["sw", "ElapsedMilliseconds"] or
/// ["System", "Threading", "Thread", "Sleep"]). Already split by the caller;
/// NOT lowercased — this function handles casing.
///
/// `ctx` provides compiler-specific callbacks for scope lookup.
///
/// Returns a `DottedResolution` telling the compiler what to emit.
pub fn resolve_dotted_name(parts: &[&str], ctx: &ResolutionContext) -> DottedResolution {
    if parts.is_empty() {
        return DottedResolution::Unresolved;
    }

    let lower_parts: Vec<String> = parts.iter().map(|p| p.to_lowercase()).collect();
    let first = &lower_parts[0];

    // ── Step 0: No-op methods ────────────────────────────────────────────
    // If the LAST part is a known no-op method, short-circuit.
    if let Some(last) = lower_parts.last() {
        if is_noop_method(last) {
            return DottedResolution::NoOp;
        }
    }

    // ── Step 1: Local variable (highest priority) ────────────────────────
    // If the first part is a local variable, everything after it is instance
    // member access. This is how .NET works: locals shadow namespaces.
    if (ctx.is_local)(first) {
        return DottedResolution::InstanceMember {
            local: first.clone(),
            members: lower_parts[1..].to_vec(),
        };
    }

    // ── Step 2: Class field (Me.field implicit) ──────────────────────────
    if (ctx.is_class_field)(first) && lower_parts.len() > 1 {
        return DottedResolution::InstanceMember {
            local: first.clone(),
            members: lower_parts[1..].to_vec(),
        };
    }

    // ── Step 3: Fully-qualified namespace (System.X.Y.Z) ────────────────
    // Try to match the longest prefix against the import list.
    // This handles both `System.Threading.Thread.Sleep` (direct FQ) and
    // static type access through fully-qualified paths.
    if let Some(res) = try_resolve_via_imports(&lower_parts, ctx.imports) {
        return res;
    }

    // ── Step 4: Imports-resolved (bare type → prepend import prefix) ─────
    // e.g. "Thread.Sleep" with import "system.threading" →
    //      try resolving "system.threading.thread.sleep"
    if !is_namespace_root(first) {
        // Only try imports for names that aren't already namespace roots
        for import_path in ctx.imports {
            let mut expanded: Vec<String> = import_path.split('.').map(|s| s.to_string()).collect();
            expanded.extend(lower_parts.iter().cloned());
            let expanded_refs: Vec<&str> = expanded.iter().map(|s| s.as_str()).collect();
            if let Some(res) = try_resolve_via_imports_refs(&expanded_refs, ctx.imports) {
                return res;
            }
        }
    }

    // ── Step 5: Known namespace root but no import match ─────────────────
    // Fall back to namespace object chain (global_get → struct_get chain).
    // This handles enum values, static properties on namespace objects, etc.
    if is_namespace_root(first) {
        return DottedResolution::NamespaceAccess {
            parts: lower_parts,
        };
    }

    // ── Step 6: Try expanding with imports for namespace access ───────────
    // e.g. "Stopwatch.StartNew" with import "system.diagnostics"
    for import_path in ctx.imports {
        let mut expanded: Vec<String> = import_path.split('.').map(|s| s.to_string()).collect();
        expanded.extend(lower_parts.iter().cloned());
        let first_expanded = &expanded[0];
        if is_namespace_root(first_expanded) {
            return DottedResolution::NamespaceAccess {
                parts: expanded,
            };
        }
    }

    // ── Step 7: User-defined type static call ────────────────────────────
    if (ctx.is_user_type)(first) {
        return DottedResolution::NamespaceAccess {
            parts: lower_parts,
        };
    }

    DottedResolution::Unresolved
}

// ─── Legacy passthrough ──────────────────────────────────────────────────────

/// Resolve a dotted .NET name to a `(module, function)` host import.
///
/// `parts` is the member-access chain split on `.`.
/// `interface_imports` is the active list of known namespace prefixes.
///
/// Returns `None` if no import prefix matches.
///
/// Kept for backward compatibility. Callers should migrate to `resolve_dotted_name()`.
pub fn resolve_interface_call(parts: &[&str], interface_imports: &[String]) -> Option<(String, String)> {
    let lower_parts: Vec<String> = parts.iter().map(|p| p.to_lowercase()).collect();
    let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
    match try_resolve_via_imports_refs(&refs, interface_imports) {
        Some(DottedResolution::HostCall { module, func }) => Some((module, func)),
        _ => None,
    }
}

// ─── Private helpers ─────────────────────────────────────────────────────────

/// Try to resolve a fully-qualified chain against the imports list.
/// Returns `HostCall` if it matches an import prefix.
fn try_resolve_via_imports(lower_parts: &[String], imports: &[String]) -> Option<DottedResolution> {
    let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
    try_resolve_via_imports_refs(&refs, imports)
}

fn try_resolve_via_imports_refs(lower_parts: &[&str], imports: &[String]) -> Option<DottedResolution> {
    if lower_parts.len() < 2 {
        return None;
    }

    // Try longest prefix first
    for prefix_len in (1..lower_parts.len()).rev() {
        let prefix = lower_parts[..prefix_len].join(".");
        if imports.contains(&prefix) {
            let func = lower_parts[prefix_len..].join(".");
            let module = namespace_to_host_module(&prefix);
            let mapped_func = map_host_func(module, &func);
            return Some(DottedResolution::HostCall {
                module: module.to_string(),
                func: mapped_func,
            });
        }
    }
    None
}
