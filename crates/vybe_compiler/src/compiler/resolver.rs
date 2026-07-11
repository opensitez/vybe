//! Common name resolver — the "one resolver" of namespaceplan.md.
//!
//! Phase 0: wraps the EXISTING Link-phase maps (`host_import_bindings`,
//! `host_const_bindings`, `host_namespace_aliases`, `host_package_roots`)
//! plus the global namespace tree (`vybe_emitter::namespaces`) behind one
//! query API implementing the plan's resolution order:
//!
//!   1. **Scope bindings** — locals/params/upvalues always shadow; a name
//!      bound in scope is NEVER a namespace reference.
//!   2. **ESM bindings** — named imports, namespace imports, package roots
//!      (profile `esm_defaults` mounts + user imports, §16.2 shadowing).
//!   3. **Global namespace tree** — fully-qualified path walk with
//!      transitive `Alias` dereference.
//!   4. `None` → not a namespace reference: the caller proceeds with
//!      ordinary member/receiver dispatch (member field names never
//!      consult free-function or namespace tables).
//!
//! Phase 0 adds NO call sites — the legacy resolution paths in
//! `calls.rs`/`expressions.rs` are untouched. Languages swap their
//! entry points to this module one phase at a time (JS → Python → PHP →
//! dotnet → rest), each gated on its suite + the resolution snapshot.

use super::Compiler;
use crate::emitter::namespaces::{self, ResolutionTarget};
use vybe_bytecode::Op;

/// What a name (or dotted chain) resolves to, in resolution order.
#[derive(Debug, Clone)]
#[allow(dead_code)] // consumed from Phase 1 (JS migration) onward
pub(crate) enum Resolution {
    /// Named ESM binding: direct `CALL_IMPORT module/func`.
    HostImport { module: String, func: String },
    /// Named ESM binding to an `ExportEntry::Value` — inlined at use-site.
    HostConst(vybe_bytecode::Value),
    /// `import * as ns from "module"` — member access under `module`.
    NamespaceAlias { module: String },
    /// Component-Model package root (`Imports System` style); the caller
    /// joins the remaining chain into a specifier under `module_root`.
    PackageRoot { module_root: String },
    /// Global namespace tree leaf (typed).
    Tree(ResolutionTarget),
    /// A registered global callable (WinForms wrapper ctor alias:
    /// `Window.Forms.TextBox` -> the `textbox` wrapper global).
    GlobalAccess { name: String },
    /// A namespace-object chain: `global_get parts[0]` + `struct_get`
    /// walk. Enum values, static props on namespace objects, and the
    /// deliberate multi-segment detours (Task.Run -> threading routing).
    NamespaceChain { parts: Vec<String> },
    /// The chain head is a scope binding (local/global/class field) --
    /// ordinary instance member access (locals shadow, always). Callers
    /// keep their special intercepts (GUI Controls.Add, Thread.Join).
    ScopedMember { local: String, members: Vec<String> },
    /// A WinForms/layout no-op method -- emit null, skip.
    NoOp,
}

impl Compiler {
    /// Resolve a bare identifier per the plan's order. `None` = not a
    /// namespace reference (locals shadow, or the name is simply unknown).
    #[allow(dead_code)] // consumed from Phase 1 (JS migration) onward
    pub(crate) fn resolve_namespace_name(&self, name: &str) -> Option<Resolution> {
        // 1. Locals shadow, always — every legacy resolver that passed
        //    `is_local: |_| false` was a bug factory (namespaceplan.md).
        if self.scope().resolve(name).is_some() || self.scope().resolve_ci(name).is_some() {
            return None;
        }
        let key = self.canon(name);

        // 2. ESM bindings (named → const → namespace → package root).
        if let Some((module, func)) = self.host_import_bindings.get(&key) {
            return Some(Resolution::HostImport {
                module: module.clone(),
                func: func.clone(),
            });
        }
        if let Some(v) = self.host_const_bindings.get(&key) {
            return Some(Resolution::HostConst(v.clone()));
        }
        if let Some(module) = self.host_namespace_aliases.get(&key) {
            return Some(Resolution::NamespaceAlias {
                module: module.clone(),
            });
        }
        if let Some(root) = self.host_package_roots.get(&key) {
            return Some(Resolution::PackageRoot {
                module_root: root.clone(),
            });
        }

        // 3. Global namespace tree (platform surfaces register their
        //    descriptor data on first query; resolution logic stays here).
        crate::platforms::dotnet::emitter::tree_register::register_namespace_tree();
        namespaces::resolve_path(&[key.as_str()]).map(Resolution::Tree)
    }

    /// Resolve a dotted chain (`["json", "dumps"]`, `["dotnet", "system",
    /// "console", "writeline"]`) per the plan's order.
    #[allow(dead_code)] // consumed from Phase 1 (JS migration) onward
    pub(crate) fn resolve_namespace_path(&self, segments: &[&str]) -> Option<Resolution> {
        let (head, rest) = segments.split_first()?;
        if rest.is_empty() {
            return self.resolve_namespace_name(head);
        }
        // 1. A scope binding on the head means the whole chain is ordinary
        //    member access, never a namespace path.
        if self.scope().resolve(head).is_some() || self.scope().resolve_ci(head).is_some() {
            return None;
        }
        let head_key = self.canon(head);

        // 2a. ESM namespace alias: `alias.field` → (module, field). Only a
        //     single member step — deeper chains under an alias are member
        //     access on the namespace object.
        if rest.len() == 1 {
            if let Some(module) = self.host_namespace_aliases.get(&head_key) {
                return Some(Resolution::HostImport {
                    module: module.clone(),
                    func: rest[0].to_string(),
                });
            }
        }

        // 2b. Package root: build the Component-Model specifier — first
        //     join `:` (package → interface), further joins `/`, last
        //     segment is the function. Module path segments canon
        //     (CM package names are lowercase by spec); the FUNCTION
        //     keeps its source spelling — the host export's true casing
        //     is authoritative and mangling it is the `matchAll` bug
        //     class.
        if let Some(root) = self.host_package_roots.get(&head_key) {
            if let Some((func, path)) = rest.split_last() {
                if !path.is_empty() {
                    let mut module = root.trim_end_matches(':').to_string();
                    module.push(':');
                    let canon_path: Vec<String> = path.iter().map(|s| self.canon(s)).collect();
                    module.push_str(&canon_path[0]);
                    for p in &canon_path[1..] {
                        module.push('/');
                        module.push_str(p);
                    }
                    return Some(Resolution::HostImport {
                        module,
                        func: func.to_string(),
                    });
                }
            }
        }

        // 3. Global namespace tree walk (platform surfaces register their
        //    descriptor data on first query). A profile tree-mount rebases
        //    the chain onto its tree path first (`System.Math.Sin` →
        //    `dotnet.system.math.sin`); rebased segments lowercase because
        //    tree keys are lowercase-canonical and the mounted surfaces
        //    (.NET CLS) resolve case-insensitively.
        crate::platforms::dotnet::emitter::tree_register::register_namespace_tree();
        if let Some(base) = self.tree_mounts.get(&head_key) {
            let mut rebased: Vec<String> = base.split('.').map(str::to_string).collect();
            rebased.extend(rest.iter().map(|s| s.to_lowercase()));
            let refs: Vec<&str> = rebased.iter().map(|s| s.as_str()).collect();
            return namespaces::resolve_path(&refs).map(Resolution::Tree);
        }
        // 3b. Ambient roots — .NET `Imports`/`using` context as data: a bare
        //     qualified chain (`Thread.Sleep`) searches under each ambient
        //     tree root in declaration order; first hit wins.
        for root in &self.ambient_tree_roots {
            let mut expanded: Vec<String> = root.split('.').map(str::to_string).collect();
            expanded.push(head.to_lowercase());
            expanded.extend(rest.iter().map(|s| s.to_lowercase()));
            let refs: Vec<&str> = expanded.iter().map(|s| s.as_str()).collect();
            if let Some(target) = namespaces::resolve_path(&refs) {
                // Ambient expansion accepts only terminal callables/values —
                // a namespace hit under an ambient root is ambiguous prefix
                // noise, never a resolution (the legacy cascade had the same
                // NamespaceAccess-skip guard on import expansion).
                if !matches!(target, namespaces::ResolutionTarget::NamespaceObject(_)) {
                    return Some(Resolution::Tree(target));
                }
            }
        }
        let canon: Vec<String> = segments.iter().map(|s| self.canon(s)).collect();
        let refs: Vec<&str> = canon.iter().map(|s| s.as_str()).collect();
        namespaces::resolve_path(&refs).map(Resolution::Tree)
    }

    /// Sorted dump of every Link-phase (name → target) binding — the
    /// resolution-snapshot harness reads this to prove later phases don't
    /// silently change what any name resolves to.
    pub fn resolution_snapshot(&self) -> Vec<String> {
        let mut lines = Vec::new();
        for (k, (m, f)) in &self.host_import_bindings {
            lines.push(format!("named      {k} -> {m}::{f}"));
        }
        for (k, v) in &self.host_const_bindings {
            lines.push(format!("const      {k} -> {v}"));
        }
        for (k, m) in &self.host_namespace_aliases {
            lines.push(format!("namespace  {k} -> {m}"));
        }
        for (k, r) in &self.host_package_roots {
            lines.push(format!("package    {k} -> {r}"));
        }
        for (k, p) in &self.tree_mounts {
            lines.push(format!("treemount  {k} -> {p}"));
        }
        for p in &self.ambient_tree_roots {
            lines.push(format!("ambient    {p}"));
        }
        lines.sort();
        lines
    }

    /// Snapshot-harness entry: run the Link phase on `module` (profile
    /// defaults + user imports), then dump the resulting bindings.
    pub fn linked_resolution_snapshot(mut self, module: &crate::ast::Module) -> Vec<String> {
        self.link(module);
        self.resolution_snapshot()
    }
}

impl Compiler {
    /// Resolve a .NET-shaped dotted chain — the dotnet platform's OLD
    /// `resolve_dotted_name` step ordering, now living in the common
    /// resolver over platform DATA (noop list, ctor aliases, namespace
    /// roots, host_map, descriptor tree) and REAL compiler scope state
    /// (the legacy `ResolutionContext` closure stubs are gone).
    ///
    /// Step order preserved verbatim from the legacy cascade: noop →
    /// scope shadow → class field → user-type bail → WinForms ctor
    /// alias → tree (full chain) → imports (direct + expansion,
    /// ranked) → namespace roots.
    pub(crate) fn resolve_dotnet_chain(&self, parts: &[String]) -> Option<Resolution> {
        use crate::platforms::dotnet::emitter as dn;
        let lower: Vec<String> = parts.iter().map(|s| s.to_lowercase()).collect();
        let first = lower.first()?.clone();

        // Step 0: no-op methods (WinForms layout) — data.
        if lower.last().is_some_and(|l| dn::is_noop_method(l)) {
            return Some(Resolution::NoOp);
        }

        let is_user_type = |name: &str| -> bool {
            self.defined_classes.contains(name)
                || self
                    .defined_classes
                    .iter()
                    .any(|c| c.eq_ignore_ascii_case(name))
        };

        // Step 1: locals/module globals shadow namespaces — real scope
        // state (the legacy `is_local` closures reimplemented this).
        let head_is_local = !is_user_type(&first)
            && (self.scopes.iter().any(|scope| {
                scope
                    .locals
                    .iter()
                    .any(|l| l.name.eq_ignore_ascii_case(&first))
            }) || self.defined_globals.contains(&first)
                || self
                    .defined_globals
                    .iter()
                    .any(|g| g.eq_ignore_ascii_case(&first)));
        if head_is_local {
            return Some(Resolution::ScopedMember {
                local: first,
                members: lower[1..].to_vec(),
            });
        }

        // Step 2: implicit `Me.field` — fields of the current class.
        let head_is_field = self
            .current_class
            .as_ref()
            .and_then(|cn| self.pending_classes.get(cn.as_str()))
            .is_some_and(|pc| pc.fields.iter().any(|f| f.eq_ignore_ascii_case(&first)));
        if head_is_field && lower.len() > 1 {
            return Some(Resolution::ScopedMember {
                local: first,
                members: lower[1..].to_vec(),
            });
        }

        // Step 2b: user-defined type — bail so static class dispatch
        // handles it (never a namespace FQN).
        if is_user_type(&first) {
            return None;
        }

        // Step 2c: WinForms constructor aliases (`Window.Forms.X`) — data.
        if lower.len() >= 2 {
            let type_name = lower.last().unwrap();
            let prefix = lower[..lower.len() - 1].join(".");
            if (prefix.eq_ignore_ascii_case("system.windows.forms")
                || prefix.eq_ignore_ascii_case("window.forms"))
                && dn::lookup_component_constructor(type_name).is_some()
            {
                return Some(Resolution::GlobalAccess {
                    name: type_name.clone(),
                });
            }
        }

        // Steps 2d/3: the chain against the descriptor tree + imports.
        let imports = {
            let mut v = crate::emitter::dotnet::surface().default_imports().to_vec();
            v.extend(self.profile.namespaces.extra_imports.clone());
            v
        };
        if let Some(r) = self.dotnet_via_imports(&lower, &imports) {
            return Some(r);
        }

        // Step 4: bare chain expanded under each ambient import, ranked
        // (callables beat namespace chains; longer imports win ties).
        let first_is_ns_root = dn::is_namespace_root(&first);
        let mut best: Option<(Resolution, usize, usize)> = None;
        for import_path in &imports {
            let mut expanded: Vec<String> = import_path.split('.').map(|p| p.to_string()).collect();
            expanded.extend(lower.iter().cloned());
            if let Some(r) = self.dotnet_via_imports(&expanded, &imports) {
                if first_is_ns_root && matches!(r, Resolution::NamespaceChain { .. }) {
                    continue;
                }
                let rank = match &r {
                    Resolution::GlobalAccess { .. } => 3,
                    Resolution::NamespaceChain { .. } => 1,
                    _ => 2,
                };
                let import_parts = import_path.split('.').count();
                let better = best
                    .as_ref()
                    .is_none_or(|(_, br, bp)| rank > *br || (rank == *br && import_parts > *bp));
                if better {
                    best = Some((r, rank, import_parts));
                }
            }
        }
        if let Some((r, _, _)) = best {
            return Some(r);
        }

        // Steps 5/6: namespace-object chains.
        if dn::is_namespace_root(&first) {
            return Some(Resolution::NamespaceChain { parts: lower });
        }
        for import_path in &imports {
            let mut expanded: Vec<String> = import_path.split('.').map(|p| p.to_string()).collect();
            expanded.extend(lower.iter().cloned());
            if dn::is_namespace_root(&expanded[0]) {
                return Some(Resolution::NamespaceChain { parts: expanded });
            }
        }
        None
    }

    /// One chain against the tree + import prefixes — the tree-backed
    /// port of the legacy `try_resolve_via_imports_refs`, arm for arm,
    /// including the deliberate multi-segment `NamespaceChain` detour
    /// (Task.Run → threading routing) and the host_map fabrication tail.
    fn dotnet_via_imports(&self, lower: &[String], imports: &[String]) -> Option<Resolution> {
        use crate::platforms::dotnet::emitter as dn;
        if lower.len() < 2 {
            return None;
        }
        if let Some(r) = self.dotnet_tree_static(lower) {
            return Some(r);
        }
        for prefix_len in (1..lower.len()).rev() {
            let prefix = lower[..prefix_len].join(".");
            if imports.iter().any(|i| i == &prefix) {
                let suffix = &lower[prefix_len..];
                if suffix.len() == 1 {
                    let mut segs: Vec<String> = lower[..prefix_len].to_vec();
                    segs.push(suffix[0].clone());
                    if let Some(r) = self.dotnet_tree_static(&segs) {
                        return Some(r);
                    }
                }
                if suffix.len() > 1 {
                    return Some(Resolution::NamespaceChain {
                        parts: lower.to_vec(),
                    });
                }
                if dn::is_namespace_root(&suffix[0]) {
                    return Some(Resolution::NamespaceChain {
                        parts: lower.to_vec(),
                    });
                }
                let func = suffix.join(".");
                let module = dn::namespace_to_host_module(&prefix);
                let mapped = dn::map_host_func(module, &func);
                return Some(Resolution::HostImport {
                    module: module.to_string(),
                    func: mapped,
                });
            }
        }
        None
    }

    /// Static member walk of the platform-registered `dotnet.*` tree.
    fn dotnet_tree_static(&self, lower: &[String]) -> Option<Resolution> {
        if lower.len() < 2 {
            return None;
        }
        crate::platforms::dotnet::emitter::tree_register::register_namespace_tree();
        let mut segs: Vec<&str> = Vec::with_capacity(lower.len() + 1);
        segs.push("dotnet");
        segs.extend(lower.iter().map(|s| s.as_str()));
        match namespaces::resolve_path(&segs)? {
            t @ (ResolutionTarget::CommonEmit(_) | ResolutionTarget::HostCall { .. }) => {
                Some(Resolution::Tree(t))
            }
            _ => None,
        }
    }
}
