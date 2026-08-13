//! Common name resolver — the "one resolver" of namespaceplan.md.
//!
//! Phase 0: wraps the EXISTING Link-phase maps (`host_import_bindings`,
//! `host_const_bindings`, `host_namespace_aliases`, `host_package_roots`)
//! plus the global namespace tree (`vybe_runtime::namespaces`) behind one
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
use crate::primitives::namespaces::{self, ResolutionTarget};

/// Lazy platform-tree registration: every platform/language package
/// contributes its descriptor DATA to the shared tree before a walk
/// (namespaceplan.md roots: `dotnet.*`, `plib.*`, `libc.*`, `php.*`,
/// `dart.*`; `ecma.*`/`wasi.*` mount from the host FunctionRegistry
/// inside `resolve_path` itself). Each registrar is Once-guarded.
pub(super) fn register_platform_trees() {
    // Platforms register THEMSELVES, from the language plugin that needs them
    // (`vybe_language_dart::register()` -> `vybe_platform_flutter::register()`),
    // which mounts the tree at plugin-registration time. The compiler no longer
    // names a platform, so a platform can become a dylib.
    crate::ensure_languages_registered();
    vybe_runtime::registry::register_all_platform_trees();
    vybe_runtime::registry::register_all_trees();
}

/// What a name (or dotted chain) resolves to, in resolution order.
#[derive(Debug, Clone)]
#[allow(dead_code)] // consumed from Phase 1 (JS migration) onward
pub(crate) enum Resolution {
    /// Named ESM binding: direct `CALL_IMPORT module/func`.
    HostImport { module: String, func: String },
    /// Named ESM binding to an `ExportEntry::Value` — inlined at use-site.
    HostConst(vybe_runtime::Value),
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
    /// A namespace-tree value leaf followed by ordinary member accesses,
    /// e.g. `System.Environment.CurrentDirectory.Length`.
    ResolvedPrefix {
        target: ResolutionTarget,
        suffix: Vec<String>,
    },
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
        if self.scope().resolve(name).is_some() {
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
        register_platform_trees();
        if let Some(target) = namespaces::resolve_path(&[key.as_str()]) {
            return Some(Resolution::Tree(target));
        }

        // 3b. Ambient roots — a bare unqualified name searches under each
        //     mounted ambient tree root in declaration order (the single-name
        //     analogue of the dotted-chain rule in `resolve_namespace_path`).
        //     PURE FALLBACK: only reached when the name resolves to nothing at
        //     top level, so it can only add resolutions, never change an
        //     existing one. Narrowed further to a `Ctor` carrying a
        //     construction `spec` — a config-object type constructed through
        //     the common-resolver path (`flutter.scaffold`). This deliberately
        //     excludes spec-less `Ctor`s (the dotnet BCL registers its tree
        //     Types with `ctor: None`, reaching construction via
        //     `lookup_constructor` + dotted chains instead), so mounting a
        //     `.NET` ambient root here never re-routes a bare `Console`.
        let ambient_key = key.to_lowercase();
        for root in &self.ambient_tree_roots {
            let mut expanded: Vec<String> = root.split('.').map(str::to_string).collect();
            expanded.push(ambient_key.clone());
            let refs: Vec<&str> = expanded.iter().map(String::as_str).collect();
            if let Some(target @ namespaces::ResolutionTarget::Ctor { spec: Some(_), .. }) =
                namespaces::resolve_path(&refs)
            {
                return Some(Resolution::Tree(target));
            }
        }

        None
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
        if self.scope().resolve(head).is_some() {
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
        register_platform_trees();
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

    /// Resolve a source-level type name against the `user.<unit>.*` root.
    ///
    /// Returns the ONE canonical identity a `NamespaceDecl` produced, which is
    /// the name every downstream consumer already keys on
    /// (`defined_classes`, `normalized_classes`, the declared-type hint).
    ///
    /// The lookup order is the plan's, applied to the user root:
    ///
    /// 1. the spelling as written — a fully-qualified `MyApp.Models.Customer`,
    /// 2. relative to the enclosing `Namespace` — a sibling type named bare,
    /// 3. under each imported namespace — `Imports MyApp.Models` + `Customer`.
    ///
    /// Locals are NOT consulted here: a type position is not a value position,
    /// and a variable named `Customer` does not shadow the type `Customer`.
    pub(crate) fn resolve_user_namespace_type(&self, name: &str) -> Option<String> {
        let normalized = namespaces::normalize_source_path(name);
        if normalized.is_empty() {
            return None;
        }
        let canon = self.canon(&normalized);
        let segments: Vec<&str> = canon.split('.').filter(|s| !s.is_empty()).collect();
        if segments.is_empty() {
            return None;
        }

        let walk = |prefix: Option<&str>| -> Option<String> {
            let mut path: Vec<&str> = Vec::new();
            if let Some(prefix) = prefix {
                path.extend(prefix.split('.').filter(|s| !s.is_empty()));
            }
            path.extend_from_slice(&segments);
            match namespaces::resolve_in(&self.user_namespace_tree, &path) {
                Some(ResolutionTarget::Ctor { path, .. }) => Some(path),
                // A NamespaceObject means the name is a namespace, not a type
                // — `MyApp.Models` in a type position resolves to nothing, the
                // same answer real compilers give.
                _ => None,
            }
        };

        if let Some(hit) = walk(None) {
            return Some(hit);
        }
        // Enclosing namespace, innermost first: inside `A.B`, `T` finds `A.B.T`
        // before `A.T` — the shadowing rule every namespaced language shares.
        // `source_namespace_contexts` supplies the enclosing `Namespace` AND
        // the namespace of the class being compiled, which are not the same
        // thing: a method body compiles with `current_namespace` already
        // unwound.
        for context in self.source_namespace_contexts() {
            let mut scope: Vec<&str> = context.split('.').filter(|s| !s.is_empty()).collect();
            while !scope.is_empty() {
                if let Some(hit) = walk(Some(&scope.join("."))) {
                    return Some(hit);
                }
                scope.pop();
            }
        }
        for prefix in &self.source_namespace_imports {
            if let Some(hit) = walk(Some(prefix)) {
                return Some(hit);
            }
        }
        None
    }

    /// A DECLARED type hint, reduced to the one canonical identity.
    ///
    /// `Dim c As MyApp.Models.Customer` and `Dim c As Customer` name one class,
    /// and every consumer of a hint keys on the identity `NamespaceDecl`
    /// lowering produced. Without this the qualified spelling matched nothing,
    /// the local silently degraded to dynamic, and its property writes went to
    /// the object's dynamic bag while the class's own methods kept reading the
    /// declared field — one property with two storages.
    ///
    /// Only a spelling that RESOLVES is rewritten, so a builtin (`Integer`), a
    /// platform type (`System.Text.StringBuilder`) and an unknown name are all
    /// passed through untouched.
    pub(crate) fn canonicalize_declared_type_hint(
        &self,
        type_hint: Option<vybe_ast::TypeHint>,
    ) -> Option<vybe_ast::TypeHint> {
        let mut type_hint = type_hint?;
        let spelling = type_hint.spelling();
        // Only a plain dotted name. A generic argument list or an array suffix
        // is a COMPOSITE spelling whose parts each need resolving; rewriting
        // the whole string against a type name would corrupt it.
        if !spelling.contains('.') || spelling.contains(['<', '[', '(', ' ', ',']) {
            return Some(type_hint);
        }
        if let Some(canonical) = self.resolve_user_namespace_type(spelling) {
            type_hint.set_spelling(canonical);
        }
        Some(type_hint)
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

#[cfg(test)]
mod user_root_tests {
    use crate::primitives::Compiler;

    /// The `user.<unit>.*` root resolves a class declared inside a
    /// `Namespace` by every spelling that names it.
    fn compiler_with(classes: &[&str]) -> Compiler {
        // A bare profile: the user root is language-neutral, and linking a
        // language crate just to get one would make this test depend on
        // whichever language happened to register first.
        let profile = crate::profile::parse_profile("[compiler]\n").expect("bare profile");
        let mut compiler = Compiler::with_profile(profile);
        for class in classes {
            compiler.defined_classes.insert((*class).to_string());
        }
        compiler.build_user_namespace_tree();
        compiler
    }

    #[test]
    fn qualified_spelling_resolves_to_the_declared_identity() {
        let c = compiler_with(&["myapp.models.customer"]);
        assert_eq!(
            c.resolve_user_namespace_type("myapp.models.customer").as_deref(),
            Some("myapp.models.customer")
        );
    }

    #[test]
    fn single_segment_namespace_resolves() {
        let c = compiler_with(&["myapp.customer"]);
        assert_eq!(
            c.resolve_user_namespace_type("myapp.customer").as_deref(),
            Some("myapp.customer")
        );
    }

    #[test]
    fn a_namespace_is_not_a_type() {
        let c = compiler_with(&["myapp.models.customer"]);
        assert_eq!(c.resolve_user_namespace_type("myapp.models"), None);
    }

    #[test]
    fn unknown_name_resolves_to_nothing() {
        let c = compiler_with(&["myapp.models.customer"]);
        assert_eq!(c.resolve_user_namespace_type("myapp.models.order"), None);
    }
}

impl Compiler {
    /// Resolve a dotted chain through the profile-mounted namespace tree.
    ///
    /// The compiler owns only the common resolution rules: scope shadowing,
    /// tree mounts, ambient roots, and namespace objects. Platform-specific
    /// surface lives in `vybe_runtime::namespaces` registrations.
    pub(crate) fn resolve_profile_namespace_chain(&self, parts: &[String]) -> Option<Resolution> {
        let first = parts.first()?;
        let lower: Vec<String> = parts.iter().map(|s| self.canon(s)).collect();

        let is_user_type = |name: &str| -> bool {
            self.defined_classes.contains(name)
                || self
                    .defined_classes
                    .iter()
                    .any(|c| c.eq_ignore_ascii_case(name))
        };

        let head_is_local = !is_user_type(&lower[0])
            && (self.scopes.iter().any(|scope| {
                scope
                    .locals
                    .iter()
                    .any(|l| l.name.eq_ignore_ascii_case(first))
            }) || self.defined_globals.contains(&lower[0])
                || self
                    .defined_globals
                    .iter()
                    .any(|g| g.eq_ignore_ascii_case(first)));
        if head_is_local {
            return Some(Resolution::ScopedMember {
                local: lower[0].clone(),
                members: lower[1..].to_vec(),
            });
        }

        let head_is_field = self
            .current_class
            .as_ref()
            .and_then(|cn| self.pending_classes.get(cn.as_str()))
            .is_some_and(|pc| pc.fields.iter().any(|f| f.eq_ignore_ascii_case(first)));
        if head_is_field && lower.len() > 1 {
            return Some(Resolution::ScopedMember {
                local: lower[0].clone(),
                members: lower[1..].to_vec(),
            });
        }

        if is_user_type(&lower[0]) {
            return None;
        }

        let refs: Vec<&str> = parts.iter().map(String::as_str).collect();
        if let Some(resolution) = self.resolve_namespace_path(&refs) {
            return Some(match resolution {
                Resolution::Tree(ResolutionTarget::NamespaceObject(path)) => {
                    Resolution::NamespaceChain {
                        parts: path.split('.').map(str::to_string).collect(),
                    }
                }
                other => other,
            });
        }

        for prefix_len in (1..parts.len()).rev() {
            let prefix_refs: Vec<&str> = parts[..prefix_len].iter().map(String::as_str).collect();
            let Some(Resolution::Tree(target)) = self.resolve_namespace_path(&prefix_refs) else {
                continue;
            };
            if matches!(
                target,
                ResolutionTarget::CommonEmit(_)
                    | ResolutionTarget::HostCall { .. }
                    | ResolutionTarget::Const(_)
            ) {
                return Some(Resolution::ResolvedPrefix {
                    target,
                    suffix: lower[prefix_len..].to_vec(),
                });
            }
        }
        None
    }
}
