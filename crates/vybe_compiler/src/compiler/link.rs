//! Module linking: import tables, type/struct predeclaration, static containers.
//!
//! Extracted from `compiler/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;

impl Compiler {
    /// Drain the compiler's host-import metadata into the shape the VM
    /// setup expects.
    pub(super) fn collected_host_imports(&self) -> HostImportMetadata {
        let mut named: Vec<HostImportNamed> = self
            .host_import_bindings
            .iter()
            .map(|(local, (module, func))| HostImportNamed {
                local: local.clone(),
                module: module.clone(),
                func: func.clone(),
            })
            .collect();
        named.sort_by(|a, b| a.local.cmp(&b.local));
        let mut wildcard: Vec<HostWildcardImport> = self
            .host_namespace_aliases
            .iter()
            .map(|(alias, module)| HostWildcardImport {
                alias: alias.clone(),
                module: module.clone(),
            })
            .collect();
        wildcard.sort_by(|a, b| a.alias.cmp(&b.alias));
        HostImportMetadata { named, wildcard }
    }

    pub(super) fn normalize_import_table(chunks: &mut [Chunk]) {
        if chunks.is_empty() {
            return;
        }

        let original_script_imports = chunks[0].imports.clone();
        let mut unified: Vec<BytecodeImport> = Vec::new();
        let mut remaps: Vec<Vec<u16>> = Vec::with_capacity(chunks.len());

        for chunk in chunks.iter() {
            let mut remap = Vec::with_capacity(chunk.imports.len());
            for imp in &chunk.imports {
                let idx = unified
                    .iter()
                    .position(|existing| existing.module == imp.module && existing.name == imp.name)
                    .unwrap_or_else(|| {
                        unified.push(imp.clone());
                        unified.len() - 1
                    });
                remap.push(idx as u16);
            }
            remaps.push(remap);
        }

        let script_remap = remaps.first().cloned().unwrap_or_default();

        for (chunk_idx, chunk) in chunks.iter_mut().enumerate() {
            let local_remap = &remaps[chunk_idx];
            let code = &mut chunk.code;
            let mut ip = 0usize;
            while ip + 3 < code.len() {
                let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
                let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
                let Some(op) = Op::decode(group, sub) else {
                    ip += 4;
                    continue;
                };

                let operand_start = ip + 4;
                let operand_len = op.operand_format().size_in(code, operand_start);
                if op == Op::CALL_IMPORT && operand_start + 1 < code.len() {
                    let old_idx =
                        u16::from_be_bytes([code[operand_start], code[operand_start + 1]]);
                    let remapped = local_remap
                        .get(old_idx as usize)
                        .copied()
                        .or_else(|| script_remap.get(old_idx as usize).copied());
                    if local_remap.get(old_idx as usize).is_none()
                        && std::env::var("VYBE_DEBUG_IMPORTS").is_ok()
                    {
                        eprintln!(
                            "[import-remap] chunk {} ip {}: CALL_IMPORT idx {} out of local table (len {}) — script fallback → {:?}",
                            chunk_idx,
                            ip,
                            old_idx,
                            local_remap.len(),
                            script_remap.get(old_idx as usize).copied()
                        );
                    }
                    if let Some(new_idx) = remapped {
                        let bytes = new_idx.to_be_bytes();
                        code[operand_start] = bytes[0];
                        code[operand_start + 1] = bytes[1];
                    }
                }

                ip = operand_start + operand_len;
            }
        }

        if unified.is_empty() && !original_script_imports.is_empty() {
            unified = original_script_imports;
        }
        chunks[0].imports = unified;
        for chunk in chunks.iter_mut().skip(1) {
            chunk.imports.clear();
        }
    }

    pub(super) fn predeclare_type_names(&mut self, body: &[Statement], namespace: Option<&str>) {
        for stmt in body {
            match &stmt.kind {
                StmtKind::NamespaceDecl { name, body } => {
                    let member = self.canon(name).replace('\\', ".");
                    let qualified = namespace
                        .map(|prefix| format!("{prefix}.{member}"))
                        .unwrap_or(member);
                    self.predeclare_type_names(body, Some(&qualified));
                }
                StmtKind::ClassDecl { name, .. } | StmtKind::StructDecl { name, .. } => {
                    let member = self.canon(name);
                    self.defined_globals.insert(member.clone());
                    self.defined_classes.insert(member.clone());
                    if let StmtKind::StructDecl { members, .. } = &stmt.kind {
                        self.predeclare_struct_surface(&member, members);
                    }
                    if let Some(prefix) = namespace {
                        let qualified = format!("{prefix}.{member}");
                        self.defined_globals.insert(qualified.clone());
                        self.defined_classes.insert(qualified);
                    }
                }
                StmtKind::ModuleDecl { name, members, .. } => {
                    let member = self.canon(name);
                    self.defined_globals.insert(member.clone());
                    self.defined_classes.insert(member.clone());
                    self.register_module_static_container(&member, members);
                    if let Some(prefix) = namespace {
                        let qualified = format!("{prefix}.{member}");
                        self.defined_globals.insert(qualified.clone());
                        self.defined_classes.insert(qualified);
                    }
                }
                _ => {}
            }
        }
    }

    pub(super) fn predeclare_function_names(&mut self, body: &[Statement]) {
        for stmt in body {
            let StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                is_generator,
                ..
            } = &stmt.kind
            else {
                continue;
            };

            let cname = self.canon(name);
            self.defined_globals.insert(cname.clone());
            self.defined_functions.insert(cname.clone());
            if *is_generator && self.profile.buffered_iterator_methods {
                self.generator_functions.insert(cname.clone());
            }
            self.function_param_modes
                .entry(cname.clone())
                .or_insert_with(|| params.iter().map(|param| param.pass_by).collect());
            self.function_param_types
                .entry(cname.clone())
                .or_insert_with(|| params.iter().map(|param| param.type_hint.clone()).collect());
            self.function_min_arity
                .entry(cname.clone())
                .or_insert_with(|| {
                    params
                        .iter()
                        .take_while(|param| param.default.is_none() && !param.is_rest)
                        .count()
                });
            self.function_signatures
                .entry(cname.clone())
                .or_default()
                .push(CallSignature::from_params(params));
            if let Some(return_type) = return_type.as_ref() {
                self.function_return_types
                    .entry(cname)
                    .or_insert_with(|| return_type.clone());
            }
        }
    }

    pub(super) fn predeclare_struct_surface(&mut self, name: &str, members: &[ClassMember]) {
        let mut fields = Vec::new();
        let mut instance_member_names = Vec::new();
        let mut instance_pointer_method_names = Vec::new();
        let mut instance_field_types = HashMap::new();
        let mut static_fields = Vec::new();
        let mut static_field_types = HashMap::new();
        let mut static_method_names = Vec::new();

        for member in members {
            match member {
                ClassMember::Field {
                    name,
                    type_hint,
                    modifiers,
                    ..
                } => {
                    let field_name = self.canon(name);
                    // `is_shared` (VB) and `is_static` (java/C#/…) both mean
                    // "static member" — either flag registers it as static.
                    if modifiers.is_shared || modifiers.is_static {
                        static_fields.push(field_name.clone());
                        if let Some(type_hint) = type_hint.as_ref() {
                            static_field_types
                                .insert(field_name, Self::normalize_type_hint(type_hint));
                        }
                    } else {
                        fields.push(field_name.clone());
                        if let Some(type_hint) = type_hint.as_ref() {
                            instance_field_types
                                .insert(field_name, Self::normalize_type_hint(type_hint));
                        }
                    }
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl {
                        name: method_name,
                        modifiers,
                        params,
                        ..
                    } = &stmt.kind
                    {
                        let canonical = self.canon(method_name);
                        if modifiers.is_shared || modifiers.is_static {
                            static_method_names.push(canonical);
                        } else {
                            if params
                                .first()
                                .and_then(|param| param.type_hint.as_deref())
                                .is_some_and(|type_hint| type_hint.trim_start().starts_with('*'))
                            {
                                instance_pointer_method_names.push(canonical.clone());
                            }
                            instance_member_names.push(canonical);
                        }
                    }
                }
                _ => {}
            }
        }

        self.defined_globals.insert(format!("{}$arity0", name));
        self.pending_classes
            .entry(name.to_string())
            .or_insert(PendingClass {
                parent: None,
                enclosing_class: self.current_class.clone(),
                fields,
                field_storage_names: HashMap::new(),
                is_value_type: true,
                instance_member_names,
                instance_pointer_method_names,
                instance_field_types,
                static_fields,
                static_field_types,
                static_method_names,
                instance_method_overloads: HashMap::new(),
                static_method_overloads: HashMap::new(),
                nested_types: Vec::new(),
                statics: Vec::new(),
            });
    }

    pub(super) fn register_module_static_container(
        &mut self,
        module_name: &str,
        members: &[ClassMember],
    ) {
        let mut module_static_fields: Vec<String> = Vec::new();
        let mut module_static_field_types: HashMap<String, String> = HashMap::new();
        let mut module_static_methods: Vec<String> = Vec::new();
        let mut module_nested_types: Vec<String> = Vec::new();

        for member in members {
            match member {
                ClassMember::Field {
                    name, type_hint, ..
                } => {
                    let field_name = self.canon(name);
                    module_static_fields.push(field_name.clone());
                    if let Some(type_hint) = type_hint.as_ref() {
                        module_static_field_types
                            .insert(field_name, Self::normalize_type_hint(type_hint));
                    }
                }
                ClassMember::Const {
                    name, type_hint, ..
                } => {
                    let const_name = self.canon(name);
                    module_static_fields.push(const_name.clone());
                    if let Some(type_hint) = type_hint.as_ref() {
                        module_static_field_types
                            .insert(const_name, Self::normalize_type_hint(type_hint));
                    }
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl {
                        name: method_name,
                        params,
                        return_type,
                        ..
                    } = &stmt.kind
                    {
                        let method_canon = self.canon(method_name);
                        module_static_methods.push(method_canon.clone());
                        self.function_param_modes
                            .entry(method_canon.clone())
                            .or_insert_with(|| params.iter().map(|param| param.pass_by).collect());
                        self.function_min_arity
                            .entry(method_canon.clone())
                            .or_insert_with(|| {
                                params
                                    .iter()
                                    .take_while(|param| param.default.is_none() && !param.is_rest)
                                    .count()
                            });
                        if let Some(return_type) = return_type.clone() {
                            self.function_return_types
                                .entry(method_canon)
                                .or_insert(return_type);
                        }
                    }
                }
                ClassMember::NestedType(stmt) => {
                    if let Some(type_name) = match &stmt.kind {
                        StmtKind::ClassDecl { name, .. }
                        | StmtKind::StructDecl { name, .. }
                        | StmtKind::EnumDecl { name, .. }
                        | StmtKind::InterfaceDecl { name, .. }
                        | StmtKind::ModuleDecl { name, .. } => Some(self.canon(name)),
                        _ => None,
                    } {
                        module_nested_types.push(type_name);
                    }
                    if let StmtKind::InterfaceDecl { name, members, .. } = &stmt.kind {
                        self.register_interface_method_signatures(name, members);
                    }
                }
                _ => {}
            }
        }

        self.pending_classes.insert(
            module_name.to_string(),
            PendingClass {
                parent: None,
                enclosing_class: self.current_class.clone(),
                fields: Vec::new(),
                field_storage_names: HashMap::new(),
                is_value_type: false,
                instance_member_names: Vec::new(),
                instance_pointer_method_names: Vec::new(),
                instance_field_types: HashMap::new(),
                static_fields: module_static_fields,
                static_field_types: module_static_field_types,
                static_method_names: module_static_methods,
                instance_method_overloads: HashMap::new(),
                static_method_overloads: HashMap::new(),
                nested_types: module_nested_types,
                statics: Vec::new(),
            },
        );
    }

    /// The Linker phase — ECMA-262 §16.2.1.5 Link adapted for Vybe.
    ///
    /// Populates the three resolver maps (`host_import_bindings`,
    /// `host_namespace_aliases`, `host_package_roots`) from two
    /// sources, in order:
    ///
    ///   1. **Profile defaults** (`profile.esm_defaults`) — the
    ///      language's ambient pre-declared imports. For JS,
    ///      `console → wasi:cli` and `Math → ecma:math`. For VB,
    ///      `System` as a `PackageRoot`.
    ///   2. **User imports** (`module.imports`) — `import { X } from
    ///      "wasi:foo"` etc. Walked last so they shadow profile
    ///      defaults on key collision (ECMA-262 §16.2 lexical bindings
    ///      override module-scope defaults).
    ///
    /// `HashMap::insert` on a duplicate key replaces the value, so
    /// walking profile-then-user gives spec-correct shadowing for
    /// free.
    ///
    /// Runs before any bytecode is emitted.
    pub(super) fn link(&mut self, module: &crate::ast::Module) {
        // Phase A.1: ambient profile defaults.
        let defaults = self.profile.esm_defaults.clone();
        for d in &defaults {
            match d {
                crate::profile::EsmDefault::Named {
                    local,
                    module: m,
                    name,
                } => {
                    let key = self.canon(local);
                    self.host_import_bindings
                        .insert(key, (m.clone(), name.clone()));
                }
                crate::profile::EsmDefault::Namespace { alias, module: m } => {
                    let key = self.canon(alias);
                    self.host_namespace_aliases.insert(key, m.clone());
                }
                crate::profile::EsmDefault::ModuleExport {
                    module: m,
                    name,
                    target_module,
                    target_name,
                } => {
                    // Mount-with-rename (namespaceplan.md): the profile
                    // declares module `m`'s export surface, so a user
                    // `from m import name` / `import { name } from "m"`
                    // binds through the SAME adapter-module path below
                    // (Phase A.2 walks `module_exports`), reconciling the
                    // source-level name with the canonical host export.
                    self.module_exports
                        .entry(m.clone())
                        .or_default()
                        .insert(name.clone(), (target_module.clone(), target_name.clone()));
                    // Also index by the HOST module the language-level name
                    // mounts to (`json` → `ecma:json`), so namespace-alias
                    // member access (`json.dumps(...)`, `import json as j;
                    // j.dumps(...)`) resolves the rename regardless of which
                    // alias the namespace is bound under.
                    if let Some(host_module) = self.host_namespace_aliases.get(&self.canon(m)) {
                        let host_module = host_module.clone();
                        self.module_exports
                            .entry(host_module)
                            .or_default()
                            .insert(name.clone(), (target_module.clone(), target_name.clone()));
                    }
                }
                crate::profile::EsmDefault::TreeMount { prefix, path } => {
                    self.tree_mounts.insert(self.canon(prefix), path.clone());
                }
                crate::profile::EsmDefault::TreeAmbient { path } => {
                    self.ambient_tree_roots.push(path.clone());
                }
                crate::profile::EsmDefault::PackageRoot {
                    prefix,
                    module_root,
                } => {
                    // Component Model package names are lowercase by
                    // spec; store + look up in lowercase regardless of
                    // the language's case sensitivity.
                    let key = prefix.to_ascii_lowercase();
                    self.host_package_roots.insert(key, module_root.clone());
                }
            }
        }

        // Phase A.2: user imports — shadow profile defaults on key
        // collision. Resolves host-specifier paths (wasi:* / wasm:* /
        // vybe:*) directly, and Adapter-module paths (node:*, etc.)
        // by walking the re-export chain in `module_exports` to the
        // ultimate target. Relative paths still resolve at bundle
        // load time.
        let bare_aliases = self.profile.bare_module_aliases.clone();
        let normalize_bare = |path: &str| -> String {
            // Profile-driven: JS routes `'fs'` → `'node:fs'` via the
            // [bare_module_aliases] table; Python's profile leaves it
            // empty so `import os` keeps Python's stdlib semantics.
            bare_aliases
                .get(path)
                .cloned()
                .unwrap_or_else(|| path.to_string())
        };
        for imp in &module.imports {
            match &imp.kind {
                crate::ast::ImportKind::Simple {
                    path,
                    alias: Some(alias),
                } => {
                    self.source_type_aliases
                        .insert(self.canon(alias), path.clone());
                    // ESM §16.2: `import X as j` rebinds X's module namespace
                    // under `j` — same binding, second name. Covers Python
                    // `import json as j` (j.dumps → ecma:json) and any
                    // language whose walker emits Simple-with-alias.
                    let path_key = self.canon(path);
                    if let Some(m) = self.host_namespace_aliases.get(&path_key).cloned() {
                        self.host_namespace_aliases.insert(self.canon(alias), m);
                    }
                }
                crate::ast::ImportKind::Named { path, names, .. } => {
                    let path = normalize_bare(path);
                    if is_host_specifier(&path) {
                        for n in names {
                            let raw_local = n.alias.as_ref().unwrap_or(&n.name).clone();
                            let key = self.canon(&raw_local);
                            // Check if this export is a constant Value (not callable).
                            // Value exports are inlined at use-site; Function exports
                            // route through CALL_IMPORT.
                            if let Some(val) = self
                                .module_value_exports
                                .get(&path)
                                .and_then(|m| m.get(&n.name))
                                .cloned()
                            {
                                self.host_const_bindings.insert(key, val);
                            } else {
                                self.host_import_bindings
                                    .insert(key, (path.clone(), n.name.clone()));
                            }
                        }
                    } else if let Some(adapter_exports) = self.module_exports.get(&path).cloned() {
                        // Adapter module: each name is a pre-resolved
                        // `(final_module, final_name)` pair courtesy
                        // of the Indirect chain walker in the Bundle.
                        for n in names {
                            let raw_local = n.alias.as_ref().unwrap_or(&n.name).clone();
                            let key = self.canon(&raw_local);
                            if let Some(target) = adapter_exports.get(&n.name).cloned() {
                                self.host_import_bindings.insert(key, target);
                            }
                            // Unresolved export — leave it; Phase 8
                            // will surface a link error here.
                        }
                    }
                    // Relative / file-system imports — bundle-level
                    // resolver handles them by inlining sources.
                }
                crate::ast::ImportKind::Wildcard { path, alias } => {
                    let path = normalize_bare(path);
                    if !is_host_specifier(&path) {
                        continue;
                    }
                    if let Some(ns) = alias {
                        let key = self.canon(ns);
                        self.host_namespace_aliases.insert(key, path);
                    }
                }
                // Simple imports (`Imports System.Text` / `using X;`) will
                // feed ambient tree roots when the legacy dotnet cascade is
                // deleted — until then bare-name resolution stays with the
                // cascade fallback (ambient duplicates shadowed the
                // compiler's Task.Run THREAD_SPAWN special path; each
                // ambient entry needs per-entry verification first).
                // Default + Simple: no meaning for host modules; skip.
                crate::ast::ImportKind::Default { .. } | crate::ast::ImportKind::Simple { .. } => {}
            }
        }

        if self.profile.name == "go" {
            for stmt in &module.body {
                let StmtKind::Expr(expr) = &stmt.kind else {
                    continue;
                };
                let ExprKind::Call { callee, args, .. } = &expr.kind else {
                    continue;
                };
                if !matches!(&callee.kind, ExprKind::Ident(name) if name == "__go_named_type")
                    || args.len() != 2
                {
                    continue;
                }
                let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind else {
                    continue;
                };
                let type_name = match &args[1].value.kind {
                    ExprKind::Lit(Literal::Str(type_name)) => Some(type_name.clone()),
                    ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
                    _ => None,
                };
                if let Some(type_name) = type_name {
                    self.source_type_aliases.insert(self.canon(name), type_name);
                }
            }
        }
    }

    pub(crate) fn resolve_source_type_alias(&self, name: &str) -> String {
        let normalized = Self::strip_global_namespace_prefix(name);
        let trimmed = normalized.trim().replace('\\', ".");
        let (head, tail) = trimmed
            .split_once('.')
            .map(|(head, tail)| (head.trim(), Some(tail.trim())))
            .unwrap_or((trimmed.as_str(), None));
        let (alias_head, suffix) = head
            .strip_suffix("()")
            .map(|bare| (bare.trim_end(), "()"))
            .unwrap_or((head, ""));
        let key = self.canon(alias_head);
        let Some(target) = self.source_type_aliases.get(&key) else {
            return trimmed;
        };
        match tail {
            Some(tail) if !tail.is_empty() => format!("{}{}.{}", target, suffix, tail),
            _ => format!("{}{}", target, suffix),
        }
    }
}
