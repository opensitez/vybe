//! Common name resolver — the "one resolver" of namespaceplan.md.
//!
//! Phase 0: wraps the EXISTING Link-phase maps (`host_import_bindings`,
//! `host_const_bindings`, `host_namespace_aliases`, `host_package_roots`)
//! plus the global namespace tree (`crate::primitives::namespaces`) behind one
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
use crate::primitives::namespaces::UserGlobalKind;

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
    // ── The typed tree query, as ONE surface ────────────────────────────
    //
    // Principle 7 of flexclassplan.md: a single choke point answers "what
    // implements operation X on receiver Y". These wrap the tree walks so a
    // call site states only the QUESTION — the scope list and the case fold are
    // the resolver's to supply, and getting either wrong at a call site is a
    // silent miss rather than an error.
    //
    // ⛔ NO SCOPE OR ESM SHADOWING HERE, deliberately. `resolve_namespace_name`
    // lets a local shadow a namespace, because a bare name is ambiguous. A
    // typed member lookup is asked only once the RECEIVER'S TYPE is settled, so
    // the shadowing question is already answered; re-asking it would let a
    // local named like a registered type change member resolution.

    fn tree_type_scope(&self) -> &[String] {
        &self.profile.namespaces.type_scopes
    }

    pub(crate) fn tree_instance_member(
        &self,
        class_name: &str,
        member: &str,
    ) -> Option<namespaces::NamespaceNode> {
        namespaces::lookup_type_instance_member(
            self.tree_type_scope(),
            class_name,
            member,
            self.tree_fold(),
        )
    }

    pub(crate) fn tree_static_member(
        &self,
        class_name: &str,
        member: &str,
    ) -> Option<namespaces::NamespaceNode> {
        namespaces::lookup_type_static_member(
            self.tree_type_scope(),
            class_name,
            member,
            self.tree_fold(),
        )
    }

    pub(crate) fn tree_member_return(&self, class_name: &str, member: &str) -> Option<String> {
        namespaces::lookup_type_member_return(
            self.tree_type_scope(),
            class_name,
            member,
            self.tree_fold(),
        )
    }

    pub(crate) fn tree_instance_target(
        &self,
        class_name: &str,
        member: &str,
        argc: u8,
    ) -> Option<crate::component_classes::InstanceMethodTarget> {
        namespaces::lookup_type_instance_target(
            self.tree_type_scope(),
            class_name,
            member,
            argc,
            self.tree_fold(),
        )
    }

    pub(crate) fn tree_ctor_target(
        &self,
        class_name: &str,
    ) -> Option<crate::component_classes::ConstructorTarget> {
        namespaces::lookup_type_ctor_target(self.tree_type_scope(), class_name, self.tree_fold())
    }

    pub(crate) fn tree_ctor_spec(&self, class_name: &str) -> Option<namespaces::CtorSpec> {
        namespaces::lookup_type_ctor_spec(self.tree_type_scope(), class_name, self.tree_fold())
    }

    pub(crate) fn tree_property_target(
        &self,
        class_name: &str,
        member: &str,
    ) -> Option<crate::component_classes::InstancePropertyTarget> {
        namespaces::lookup_type_property_target(
            self.tree_type_scope(),
            class_name,
            member,
            self.tree_fold(),
        )
    }

    pub(crate) fn tree_property_setter_target(
        &self,
        class_name: &str,
        member: &str,
    ) -> Option<crate::component_classes::InstancePropertyTarget> {
        namespaces::lookup_type_property_setter_target(
            self.tree_type_scope(),
            class_name,
            member,
            self.tree_fold(),
        )
    }

    /// Does the profile's RUNTIME COLLECTION scope declare `member` at `arity`?
    ///
    /// A different scope from the typed lookups above — this asks the ambient
    /// collection surface, not a named receiver's type — so it takes its scope
    /// from `runtime_collection_scope` rather than `type_scopes`.
    pub(crate) fn tree_collection_declares(&self, member: &str, arity: u8) -> bool {
        let scope: Vec<&str> = self
            .profile
            .namespaces
            .runtime_collection_scope
            .iter()
            .map(String::as_str)
            .collect();
        namespaces::scope_declares_member_arity(&scope, member, arity, self.tree_fold())
    }

    pub(crate) fn tree_is_registered_type(&self, class_name: &str) -> bool {
        namespaces::is_registered_type(self.tree_type_scope(), class_name, self.tree_fold())
    }

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
        if let Some(target) = namespaces::resolve_path(&[key.as_str()], self.tree_fold()) {
            return Some(Resolution::Tree(target));
        }

        // 3b. Ambient roots — a bare unqualified name searches under each
        //     mounted ambient tree root in declaration order (the single-name
        //     analogue of the dotted-chain rule in `resolve_namespace_path`).
        //     PURE FALLBACK: only reached when the name resolves to nothing at
        //     top level, so it can only add resolutions, never change an
        //     existing one. Accept only terminal callables/values plus a
        //     `Ctor` carrying a construction `spec` — a config-object type
        //     constructed through the common-resolver path (`flutter.scaffold`).
        //     This deliberately excludes namespace objects and spec-less
        //     `Ctor`s (the dotnet BCL registers its tree Types with
        //     `ctor: None`, reaching construction via `lookup_constructor` +
        //     dotted chains instead), so mounting a `.NET` ambient root here
        //     never re-routes a bare `Console`.
        // Same rule as the rebase above: the tree folds at LOOKUP now, so a
        // key handed to it keeps the spelling the source wrote.
        let ambient_key = key.to_string();
        for root in &self.ambient_tree_roots {
            let mut expanded: Vec<String> = root.split('.').map(str::to_string).collect();
            expanded.push(ambient_key.clone());
            let refs: Vec<&str> = expanded.iter().map(String::as_str).collect();
            if let Some(target) = namespaces::resolve_path(&refs, self.tree_fold()) {
                match target {
                    namespaces::ResolutionTarget::HostCall { .. }
                    | namespaces::ResolutionTarget::CommonEmit(_)
                    | namespaces::ResolutionTarget::Const(_)
                    | namespaces::ResolutionTarget::Ctor { spec: Some(_), .. } => {
                        return Some(Resolution::Tree(target));
                    }
                    _ => {}
                }
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
        //    `dotnet.System.Math.Sin`).
        //
        //    ⛔ Rebased segments used to be LOWERCASED here, "because tree
        //    keys are lowercase-canonical". That premise is gone: the tree
        //    keeps the declared spelling and `fold_get` matches exact first,
        //    folding only on a miss. Folding HERE instead destroys the
        //    declared spelling before the tree ever sees it — and it does so
        //    unconditionally, for every language, so a case-sensitive one
        //    cannot opt out and a conditional fold downstream cannot repair
        //    it. Pass the source's own segments and let the lookup decide.
        register_platform_trees();
        if let Some(base) = self.tree_mounts.get(&head_key) {
            let mut rebased: Vec<String> = base.split('.').map(str::to_string).collect();
            rebased.extend(rest.iter().map(|s| s.to_string()));
            let refs: Vec<&str> = rebased.iter().map(|s| s.as_str()).collect();
            return namespaces::resolve_path(&refs, self.tree_fold()).map(Resolution::Tree);
        }
        // 3b. Ambient roots — .NET `Imports`/`using` context as data: a bare
        //     qualified chain (`Thread.Sleep`) searches under each ambient
        //     tree root in declaration order; first hit wins.
        for root in &self.ambient_tree_roots {
            let mut expanded: Vec<String> = root.split('.').map(str::to_string).collect();
            expanded.push(head.to_string());
            expanded.extend(rest.iter().map(|s| s.to_string()));
            let refs: Vec<&str> = expanded.iter().map(|s| s.as_str()).collect();
            if let Some(target) = namespaces::resolve_path(&refs, self.tree_fold()) {
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
        namespaces::resolve_path(&refs, self.tree_fold()).map(Resolution::Tree)
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
        self.resolve_user_namespace_member(name, UserGlobalKind::Type)
    }

    /// The same three tiers for a declared FUNCTION.
    ///
    /// Functions never had a tree at all: `resolve_namespaced_function_identity`
    /// ran a flat cascade over `defined_functions` ending in a unique-suffix
    /// GUESS. Types and functions are now the same question asked of the same
    /// root, discriminated by [`UserGlobalKind`] so a type position cannot
    /// accept a function or the reverse.
    pub(crate) fn resolve_user_namespace_function(&self, name: &str) -> Option<String> {
        self.resolve_user_namespace_member(name, UserGlobalKind::Function)
    }

    fn resolve_user_namespace_member(&self, name: &str, want: UserGlobalKind) -> Option<String> {
        let normalized = namespaces::normalize_source_path(name);
        if normalized.is_empty() {
            return None;
        }
        let canon = self.canon(&normalized);
        let segments: Vec<&str> = canon.split('.').filter(|s| !s.is_empty()).collect();
        if segments.is_empty() {
            return None;
        }

        let tree = self.user_namespace_tree.borrow();
        let walk = |prefix: Option<&str>| -> Option<String> {
            let mut path: Vec<&str> = Vec::new();
            if let Some(prefix) = prefix {
                path.extend(prefix.split('.').filter(|s| !s.is_empty()));
            }
            path.extend_from_slice(&segments);
            match namespaces::resolve_in(&tree, &path, self.tree_fold()) {
                Some(ResolutionTarget::UserGlobal { identity, kind }) if kind == want => {
                    Some(identity)
                }
                // A NamespaceObject means the name is a namespace, not a
                // declaration — `MyApp.Models` in a type position resolves to
                // nothing, the same answer real compilers give. A declaration
                // of the OTHER kind is equally not an answer.
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
        // An import written INSIDE a namespace names a namespace relative to it.
        //
        //     namespace App.Deep { class Helper {} }
        //     namespace App { using Deep; ... Helper.Tag() ... }
        //
        // `using Deep;` there means `App.Deep`, and real C# resolves it — the
        // program above prints `App.Deep.Helper` under the .NET SDK. Vybe died
        // with "undefined is not callable": the tiers above try the import as
        // WRITTEN (`deep.helper`) and the enclosing context against the BARE
        // name (`app.helper`), and the answer is neither — it is the two
        // combined. VB's `Imports` inside a `Namespace` has the same rule, so
        // this is one walk in the common resolver, not a C# arm.
        //
        // Last, and therefore purely additive: every name that resolved before
        // still resolves to the same identity by an earlier tier, and only
        // names that previously resolved to NOTHING can reach here.
        //
        // KNOWN LIMIT: C# searches the enclosing namespaces outward for the
        // import's own name, so with both a global `Deep` and an `App.Deep` in
        // scope it binds `App.Deep`, where this binds the global one at the
        // tier above. Closing that needs each import to carry the namespace it
        // was written in; `source_namespace_imports` is a flat list filled at
        // link time and has already lost it. Not silently wrong for the case
        // above — wrong only when BOTH spellings exist, which no test covers.
        for context in self.source_namespace_contexts() {
            let mut scope: Vec<&str> = context.split('.').filter(|s| !s.is_empty()).collect();
            while !scope.is_empty() {
                let base = scope.join(".");
                for prefix in &self.source_namespace_imports {
                    if let Some(hit) = walk(Some(&format!("{base}.{prefix}"))) {
                        return Some(hit);
                    }
                }
                scope.pop();
            }
        }
        None
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
            // The real declaration entry point, not a raw set insert: the tree
            // is the storage now, so poking `defined_classes` behind its back
            // would test a state the compiler can never actually be in.
            compiler.declare_class_identity(class);
        }
        compiler
    }

    fn compiler_with_functions(functions: &[&str]) -> Compiler {
        let profile = crate::profile::parse_profile("[compiler]\n").expect("bare profile");
        let mut compiler = Compiler::with_profile(profile);
        for function in functions {
            compiler.declare_function_identity(function);
        }
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

    /// A class whose name repeats the namespace ROOT — `namespace Demo.Sub {
    /// class Demo }`. The root segment is a namespace AND the leaf is a type;
    /// both spellings have to keep working.
    #[test]
    fn class_named_after_its_namespace_root() {
        let c = compiler_with(&["Demo.Sub.Demo"]);
        assert_eq!(
            c.resolve_user_namespace_type("Demo.Sub.Demo").as_deref(),
            Some("Demo.Sub.Demo")
        );
        assert_eq!(c.resolve_user_namespace_type("Demo"), None);
    }

    /// The shape that was actually order-dependent: a BARE class whose name is
    /// also a namespace root. Built from a `HashSet`, whichever arrived first
    /// decided whether the root was a type or a namespace, and the other
    /// spelling stopped resolving — differently from run to run. Insertion is
    /// sorted shallowest-first so the bare type exists before anything descends
    /// through its statics.
    #[test]
    fn bare_class_and_namespace_sharing_a_root_both_resolve() {
        for order in [
            ["Demo", "Demo.Sub.Demo"],
            ["Demo.Sub.Demo", "Demo"],
        ] {
            let c = compiler_with(&order);
            assert_eq!(
                c.resolve_user_namespace_type("Demo").as_deref(),
                Some("Demo"),
                "bare class lost, insertion order {order:?}"
            );
            assert_eq!(
                c.resolve_user_namespace_type("Demo.Sub.Demo").as_deref(),
                Some("Demo.Sub.Demo"),
                "qualified class lost, insertion order {order:?}"
            );
        }
    }

    #[test]
    fn unknown_name_resolves_to_nothing() {
        let c = compiler_with(&["myapp.models.customer"]);
        assert_eq!(c.resolve_user_namespace_type("myapp.models.order"), None);
    }

    /// A declared FUNCTION resolves through the same root as a type.
    #[test]
    fn qualified_function_resolves_to_its_identity() {
        let c = compiler_with_functions(&["myapp.utils.repeat"]);
        assert_eq!(
            c.resolve_user_namespace_function("myapp.utils.repeat")
                .as_deref(),
            Some("myapp.utils.repeat")
        );
    }

    /// The kinds do not answer for each other. This is what the flat sets could
    /// never express: `defined_functions` and `defined_classes` were two
    /// unrelated sets of strings, so nothing stopped a type position from
    /// accepting a function name that happened to match.
    #[test]
    fn a_function_is_not_a_type_and_a_type_is_not_a_function() {
        let f = compiler_with_functions(&["myapp.utils.repeat"]);
        assert_eq!(f.resolve_user_namespace_type("myapp.utils.repeat"), None);

        let t = compiler_with(&["myapp.models.customer"]);
        assert_eq!(
            t.resolve_user_namespace_function("myapp.models.customer"),
            None
        );
    }

    /// A bare name whose namespace was never imported resolves to NOTHING.
    ///
    /// This is the behaviour the retired unique-suffix guess used to fake: it
    /// answered `customer` with `myapp.models.customer` purely because that was
    /// the only class whose name ended that way. Real VB/C# reject it (BC30002).
    #[test]
    fn bare_name_without_an_import_does_not_resolve() {
        let c = compiler_with(&["myapp.models.customer"]);
        assert_eq!(c.resolve_user_namespace_type("customer"), None);
    }
}

impl Compiler {
    /// Resolve a dotted chain through the profile-mounted namespace tree.
    ///
    /// The compiler owns only the common resolution rules: scope shadowing,
    /// tree mounts, ambient roots, and namespace objects. Platform-specific
    /// surface lives in `crate::primitives::namespaces` registrations.
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
