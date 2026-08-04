//! Module linking: import tables, type/struct predeclaration, static containers.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;
use vybe_ast::class_normalize::{PlatformBaseSpec, PlatformFieldGui};

fn mounted_ambient_tree_root(tree_mounts: &HashMap<String, String>, path: &str) -> Option<String> {
    let trimmed = path.trim();
    let mut segments = trimmed
        .split(['.', '\\'])
        .map(str::trim)
        .filter(|part| !part.is_empty());
    let head = segments.next()?;
    let base = tree_mounts.get(&head.to_ascii_lowercase())?;
    let mut out = base.clone();
    for segment in segments {
        out.push('.');
        out.push_str(&segment.to_ascii_lowercase());
    }
    Some(out)
}

fn platform_base_spec_from_ctor_spec(
    spec: crate::primitives::namespaces::CtorSpec,
) -> PlatformBaseSpec {
    PlatformBaseSpec {
        params: spec.params,
        fields: spec.fields,
        ancestry: spec.ancestry,
        control_fn: spec.control_fn,
        field_gui: spec
            .field_gui
            .into_iter()
            .map(|field| match field {
                crate::primitives::namespaces::FieldGui::NestOrProp(key) => {
                    PlatformFieldGui::NestOrProp(key)
                }
                crate::primitives::namespaces::FieldGui::Children => PlatformFieldGui::Children,
                crate::primitives::namespaces::FieldGui::Event(name) => {
                    PlatformFieldGui::Event(name)
                }
                crate::primitives::namespaces::FieldGui::Caption => PlatformFieldGui::Caption })
            .collect(),
        value_equality: spec.value_equality }
}

impl Compiler {
    fn bind_namespace_import_target(
        &mut self,
        key: String,
        target: crate::primitives::namespaces::ResolutionTarget,
    ) {
        match target {
            crate::primitives::namespaces::ResolutionTarget::HostCall { module, func, .. } => {
                self.host_import_bindings.insert(key, (module, func));
            }
            crate::primitives::namespaces::ResolutionTarget::Const(value) => {
                self.host_const_bindings.insert(key, value);
            }
            crate::primitives::namespaces::ResolutionTarget::CommonEmit(_)
            | crate::primitives::namespaces::ResolutionTarget::Ctor { .. }
            | crate::primitives::namespaces::ResolutionTarget::NamespaceObject(_) => {
                self.namespace_import_bindings.insert(key, target);
            }
        }
    }

    /// Drain the compiler's host-import metadata into the shape the VM
    /// setup expects.
    pub(super) fn collected_host_imports(&self) -> HostImportMetadata {
        let mut named: Vec<HostImportNamed> = self
            .host_import_bindings
            .iter()
            .map(|(local, (module, func))| HostImportNamed {
                local: local.clone(),
                module: module.clone(),
                func: func.clone() })
            .collect();
        named.sort_by(|a, b| a.local.cmp(&b.local));
        let mut wildcard: Vec<HostWildcardImport> = self
            .host_namespace_aliases
            .iter()
            .map(|(alias, module)| HostWildcardImport {
                alias: alias.clone(),
                module: module.clone() })
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

    /// Record every module-level VARIABLE name, for `primitives/globals.rs`.
    ///
    /// Only top-level statements are walked: a name assigned inside a function
    /// is a local, not a member of the global namespace.
    pub(super) fn collect_module_variable_names(&mut self, body: &[Statement]) {
        for stmt in body {
            match &stmt.kind {
                // The error boundary (`errors::wrap_module_in_error_boundary`)
                // nests the whole module body in a `Try`. It is a reporting
                // wrapper, not a scope — module-level names inside it are still
                // module-level, so descend. Without this the wrap hid every
                // top-level variable and `$GLOBALS[$k]` went blank again.
                StmtKind::Try {
                    body: try_body,
                    finally,
                    ..
                } => {
                    self.collect_module_variable_names(try_body);
                    if let Some(finally_body) = finally {
                        self.collect_module_variable_names(finally_body);
                    }
                }
                StmtKind::VarDecl { declarations, .. } => {
                    for decl in declarations {
                        if let BindingPattern::Ident(name) = &decl.pattern {
                            self.module_variable_names.insert(name.clone());
                        }
                    }
                }
                StmtKind::Assign { targets, .. } => {
                    for target in targets {
                        match &target.kind {
                            ExprKind::Ident(name) => {
                                self.module_variable_names.insert(name.clone());
                            }
                            // `globals()['y'] = 7` / `_G['y'] = 7` DECLARE `y`.
                            // The key is a literal, so the name is known here
                            // even though the binding never appears as an
                            // ordinary assignment — without this the write had
                            // no member to match and silently did nothing.
                            ExprKind::Index { object, index, .. }
                                if self.expr_is_global_namespace(object) =>
                            {
                                if let ExprKind::Lit(Literal::Str(key)) = &index.kind {
                                    self.module_variable_names.insert(key.clone());
                                }
                            }
                            _ => {}
                        }
                    }
                }
                StmtKind::Expr(expr) => {
                    // PHP desugars `$x = 5;` to an expression-statement Assign.
                    if let ExprKind::Assign { target, .. } = &expr.kind {
                        if let ExprKind::Ident(name) = &target.kind {
                            self.module_variable_names.insert(name.clone());
                        }
                    }
                }
                _ => {}
            }
        }
    }

    pub(super) fn predeclare_type_names(&mut self, body: &[Statement], namespace: Option<&str>) {
        fn namespace_member_name(
            compiler: &Compiler,
            name: &str,
            namespace: Option<&str>,
        ) -> String {
            let member = compiler.canon(name).replace('\\', ".");
            match namespace {
                Some(prefix) if !prefix.is_empty() && !member.contains('.') => {
                    format!("{prefix}.{member}")
                }
                _ => member }
        }

        for stmt in body {
            match &stmt.kind {
                StmtKind::NamespaceDecl { name, body } => {
                    let member = self.canon(name).replace('\\', ".");
                    let qualified = match namespace {
                        Some(prefix) if !prefix.is_empty() && !member.is_empty() => {
                            format!("{prefix}.{member}")
                        }
                        Some(prefix) if !prefix.is_empty() => prefix.to_string(),
                        _ => member };
                    let namespace = if qualified.is_empty() {
                        None
                    } else {
                        Some(qualified.as_str())
                    };
                    self.predeclare_type_names(body, namespace);
                }
                StmtKind::ClassDecl { name, .. } | StmtKind::StructDecl { name, .. } => {
                    let member = namespace_member_name(self, name, namespace);
                    self.defined_globals.insert(member.clone());
                    self.defined_classes.insert(member.clone());
                    if let StmtKind::StructDecl { members, .. } = &stmt.kind {
                        self.predeclare_struct_surface(&member, members);
                        // Normalize structs in the DECLARATION pass too, not
                        // only classes. Without this a struct never enters
                        // `normalized_classes`, so it is invisible to the
                        // augmentation fold (its source type resolves to
                        // nothing) and to any use site that compiles before its
                        // declaration — a Go method body calling a method on a
                        // type declared further down resolved to `undefined`.
                        //
                        // MERGED, not inserted: a type's members can arrive in
                        // several declarations. Go writes methods outside the
                        // type, so its walker emits one `StructDecl` per
                        // method; overwriting would leave the type holding only
                        // its last one.
                        if let Ok(nc) =
                            crate::primitives::class_normalize::emit::normalize_class_from_ast(
                                self,
                                stmt.span.clone(),
                                &member,
                                &[],
                                &[],
                                members,
                                &vybe_ast::ClassModifiers::default(),
                                true,
                            )
                        {
                            for special in &nc.special_methods {
                                match special.kind {
                                    vybe_ast::ProtocolSlot::GetItem => {
                                        self.classes_with_indexer.insert(member.clone());
                                    }
                                    vybe_ast::ProtocolSlot::SetItem => {
                                        self.classes_with_index_setter.insert(member.clone());
                                    }
                                    vybe_ast::ProtocolSlot::GetAttr => {
                                        self.program_has_getattr = true;
                                    }
                                    _ => {}
                                }
                            }
                            match self.normalized_classes.get_mut(&member) {
                                Some(existing) => existing.merge_partial(nc),
                                None => {
                                    self.normalized_classes.insert(member.clone(), nc);
                                }
                            }
                        }
                    }
                    // An index operator has to be known before ANY use site
                    // compiles, not when the class body does — a caller can be
                    // compiled first, and `x[i]` resolves the indexer from the
                    // receiver's static type.
                    if let StmtKind::ClassDecl {
                        parents,
                        interfaces: class_interfaces,
                        members,
                        modifiers,
                        ..
                    } = &stmt.kind
                    {
                        // Declaration pass: NORMALIZE the class once, here,
                        // and keep it. Normalization used to happen during
                        // code generation, one class at a time, so a class's
                        // member set depended on compilation order and an
                        // augmenting type (trait / mixin / promoted field)
                        // could not be looked up by name at all.
                        if let Ok(nc) =
                            crate::primitives::class_normalize::emit::normalize_class_from_ast(
                                self,
                                stmt.span.clone(),
                                &member,
                                parents,
                                class_interfaces,
                                members,
                                modifiers,
                                false,
                            )
                        {
                            // Index operators are read off the normalized
                            // class's ROLES, not off member spellings: Ruby
                            // `[]`, Dart `operator[]`, PHP `offsetGet` and
                            // Python `__getitem__` are one role under four
                            // names, and only normalization knows which
                            // language's names these are.
                            for special in &nc.special_methods {
                                match special.kind {
                                    vybe_ast::ProtocolSlot::GetItem => {
                                        self.classes_with_indexer.insert(member.clone());
                                    }
                                    vybe_ast::ProtocolSlot::SetItem => {
                                        self.classes_with_index_setter.insert(member.clone());
                                    }
                                    vybe_ast::ProtocolSlot::GetAttr => {
                                        self.program_has_getattr = true;
                                    }
                                    _ => {}
                                }
                            }
                            self.normalized_classes.insert(member.clone(), nc);
                        }
                        // NOTE: the member surface is registered later, by
                        // `predeclare_class_surfaces`, because augmentations
                        // (traits / mixins / promoted fields) must be folded in
                        // first — otherwise contributed members are missing from
                        // the registration and the order-dependence bug returns
                        // by another route. See flexclassplan.md §4c.
                    }
                    if let StmtKind::ClassDecl { members, .. } = &stmt.kind {
                        // A class's own method shadows a same-named builtin
                        // value method (`obj.add(x)` is the class's `add`, not
                        // a list's). The call site can compile before the class
                        // body, so the names have to be known here — by the
                        // time `compile_normal_class` registers them, an
                        // earlier caller has already been hijacked.
                        for m in members {
                            if let ClassMember::Method(stmt) = m {
                                if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                                    self.defined_class_methods.insert(self.canon(name));
                                }
                            }
                        }
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
                        for module_member in members {
                            let ClassMember::Method(stmt) = module_member else {
                                continue;
                            };
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
                            let function_name = self.canon(&format!("{prefix}.{member}.{name}"));
                            self.defined_globals.insert(function_name.clone());
                            self.defined_functions.insert(function_name.clone());
                            if *is_generator && self.profile.buffered_iterator_methods {
                                self.generator_functions.insert(function_name.clone());
                            }
                            self.function_param_modes
                                .entry(function_name.clone())
                                .or_insert_with(|| {
                                    params.iter().map(|param| param.pass_by).collect()
                                });
                            self.function_param_types
                                .entry(function_name.clone())
                                .or_insert_with(|| {
                                    params.iter().map(|param| param.type_hint.as_deref().map(str::to_string)).collect()
                                });
                            self.function_min_arity
                                .entry(function_name.clone())
                                .or_insert_with(|| {
                                    params
                                        .iter()
                                        .take_while(|param| {
                                            param.default.is_none() && !param.is_rest
                                        })
                                        .count()
                                });
                            self.function_signatures
                                .entry(function_name.clone())
                                .or_default()
                                .push(CallSignature::from_params(params));
                            if let Some(return_type) = return_type.as_ref() {
                                self.function_return_types
                                    .entry(function_name)
                                    .or_insert_with(|| return_type.clone());
                            }
                        }
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
                .or_insert_with(|| params.iter().map(|param| param.type_hint.as_deref().map(str::to_string)).collect());
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

    /// Register a CLASS in `pending_classes` during the declaration pass —
    /// before any body is compiled.
    ///
    /// `pending_classes` was previously filled only by `compile_normal_class`,
    /// i.e. *while generating code*, so a call site compiled earlier saw an
    /// empty table: `main`'s body saw no classes at all, and a method saw only
    /// the classes declared above it (forward references were invisible).
    /// Receiver-typed resolution therefore depended on compilation order, and a
    /// user method lost to a same-named profile value-method (`K().length()`)
    /// purely by position. `predeclare_struct_surface` already did this for
    /// structs; classes had no equivalent, and `defined_class_methods`
    /// (a flat, class-less name set) was the workaround.
    ///
    /// `compile_normal_class` still `insert`s the complete entry later, which
    /// overwrites this one — this is the *declaration*, that is the definition.
    /// Phase 2 of the declaration pass: fold every class's declared
    /// augmentations (PHP traits, Dart mixins, Ruby include/prepend, Java
    /// defaults, Go promotion) into its normalized member set.
    ///
    /// Runs after ALL classes are normalized — an augmenting type may be
    /// declared after its user — and before any member surface is registered.
    /// A language that declares no augmentations is untouched, so languages
    /// migrate one at a time with no flag day.
    /// Classify every normalized class's parent: a user class (compiled here)
    /// or a registered PLATFORM type (data in the namespace tree). Only this
    /// pass can tell them apart — the syntax is identical, and the tree is the
    /// only thing that knows `TForm`/`StatelessWidget`/`Form` are declared
    /// specs rather than compiled classes.
    ///
    /// Recording the spec here is what stops the emitter reaching for a
    /// constructor global that never existed. See flexclassplan.md §4c.
    pub(super) fn record_platform_bases(&mut self) {
        let scope = self.profile.namespaces.type_scopes.clone();
        if scope.is_empty() {
            return;
        }
        let names: Vec<String> = self.normalized_classes.keys().cloned().collect();
        for name in names {
            let Some(parent) = self
                .normalized_classes
                .get(&name)
                .and_then(|nc| nc.parent.clone())
            else {
                continue;
            };
            // A user class wins: a program may legitimately declare a class
            // whose name collides with a platform type, and its own definition
            // is the one in scope.
            if self.normalized_classes.contains_key(&self.canon(&parent))
                || self.defined_classes.contains(&self.canon(&parent))
            {
                continue;
            }
            if let Some(spec) = vybe_runtime::namespaces::lookup_type_ctor_spec(&scope, &parent) {
                if let Some(nc) = self.normalized_classes.get_mut(&name) {
                    nc.platform_base = Some(platform_base_spec_from_ctor_spec(spec));
                }
            }
        }
    }

    pub(super) fn apply_class_augmentations(&mut self) -> Result<(), String> {
        if self
            .normalized_classes
            .values()
            .all(|nc| nc.augmentations.is_empty())
        {
            return Ok(());
        }
        // Dependency order: a class must be folded AFTER every type it draws
        // from, or it copies a pre-augmentation snapshot. PHP traits may use
        // traits and Dart mixins may apply to mixins, so `trait A { use B; }`
        // + `class C { use A; }` must reach B's members through A. Iterating
        // the map in hash order would silently drop them.
        //
        // Cycles cannot be ordered; those entries are folded last, with
        // whatever their sources hold at that point. A cyclic `use` is an
        // error in every language concerned, and the languages reject it
        // before reaching here.
        let order = self.augmentation_fold_order();
        let mut available = self.normalized_classes.clone();
        for name in order {
            let Some(mut nc) = self.normalized_classes.get(&name).cloned() else {
                continue;
            };
            // Bind each augmentation to the class it actually names BEFORE
            // folding. `available` is keyed by `canon(name)` — lowercased for a
            // case-insensitive language, fully qualified for a namespaced one —
            // while `aug.from` carries the source spelling, so an exact lookup
            // inside the fold misses every PHP trait.
            for aug in &mut nc.augmentations {
                if let Some(resolved) = self.resolve_augmentation_source(&aug.from) {
                    aug.from = resolved;
                }
            }
            let mut errors = super::class_augmentation::apply_augmentations(&mut nc, &available);
            // `NextInOrder` augmentations contribute no members; they become
            // link classes spliced into the parent chain, so that `super`
            // inside a contributed method resolves through the ordinary
            // prototype hop (flexclassplan.md §4c-R). They are real classes:
            // registered here so `predeclare_class_surfaces` — which runs next
            // and iterates this very map — records their member surface too.
            let (links, chain_errors) =
                super::class_augmentation::build_super_chain(&mut nc, &available);
            errors.extend(chain_errors);
            for link in links {
                let key = self.canon(&link.name);
                // A link is a class like any other, and the class it stands
                // under names it as its parent. `parent_ctor_is_bound` is
                // answered from `defined_classes`, which the declaration pass
                // fills from AST `ClassDecl`s — and a link has none. Without
                // this the using class emits no super() call at all, so the
                // real base's field initializers never run and every inherited
                // field reads `undefined`.
                self.defined_classes.insert(key.clone());
                available.insert(key.clone(), link.clone());
                self.normalized_classes.insert(key, link);
            }
            if let Some(first) = errors.first() {
                // Go equal-depth promotion and Java default-method diamonds are
                // errors in the SOURCE language. Reported, never silently
                // resolved — flexclassplan.md §2f. A last-one-wins fold is what
                // hides them today.
                return Err(format!("augmentation conflict — {first}"));
            }
            available.insert(name.clone(), nc.clone());
            self.normalized_classes.insert(name, nc);
        }
        Ok(())
    }

    /// The `normalized_classes` key an augmentation's source name refers to.
    ///
    /// A source names its augmenting type however the language spells it —
    /// `use Timestamped;` — while the class map is keyed canonically, and for a
    /// namespaced language fully qualified (`app.traits.timestamped`). Resolve
    /// exactly first, then by an UNAMBIGUOUS `.suffix` match, which covers both
    /// same-namespace use and an imported `use App\Traits\X;` without a second
    /// alias table. Two candidates means the reference is ambiguous, and
    /// guessing one would silently pick a type the program never named.
    fn resolve_augmentation_source(&self, from: &str) -> Option<String> {
        let canon = self.canon(from);
        if self.normalized_classes.contains_key(&canon) {
            return Some(canon);
        }
        let dotted = format!(".{canon}");
        let mut matches = self
            .normalized_classes
            .keys()
            .filter(|key| key.ends_with(&dotted));
        match (matches.next(), matches.next()) {
            (Some(key), None) => Some(key.clone()),
            _ => None }
    }

    /// Classes ordered so that every augmenting type is folded before the
    /// classes that draw from it (topological over `augmentations.from`).
    /// Entries in a cycle come last — a cyclic `use`/`with` is an error in
    /// every language concerned and is rejected upstream.
    fn augmentation_fold_order(&self) -> Vec<String> {
        let mut ordered: Vec<String> = Vec::with_capacity(self.normalized_classes.len());
        let mut placed: HashSet<String> = HashSet::new();
        // Repeat until a full sweep places nothing new: anything still missing
        // is in a cycle, and is appended as-is.
        loop {
            let mut progressed = false;
            for (name, nc) in &self.normalized_classes {
                if placed.contains(name) {
                    continue;
                }
                let ready = nc.augmentations.iter().all(|aug| {
                    // Resolved the same way the fold will resolve it. Testing
                    // the RAW name here reports "ready" for every augmentation
                    // whose spelling differs from its key — which is all of
                    // them in a case-insensitive or namespaced language — so
                    // every class places on the first sweep, in hash order, and
                    // a trait that uses a trait folds a pre-augmentation
                    // snapshot. That is the exact bug this ordering exists to
                    // prevent.
                    match self.resolve_augmentation_source(&aug.from) {
                        Some(key) => placed.contains(&key),
                        None => true }
                });
                if ready {
                    ordered.push(name.clone());
                    placed.insert(name.clone());
                    progressed = true;
                }
            }
            if !progressed {
                break;
            }
        }
        for name in self.normalized_classes.keys() {
            if !placed.contains(name) {
                ordered.push(name.clone());
            }
        }
        ordered
    }

    /// Phase 3 of the declaration pass: register every class's member surface,
    /// now that augmentations have been folded in.
    pub(super) fn predeclare_class_surfaces(&mut self) {
        let entries: Vec<(String, Vec<String>)> = self
            .normalized_classes
            .iter()
            .map(|(name, nc)| (name.clone(), nc.bases.clone()))
            .collect();
        for (name, bases) in entries {
            self.predeclare_class_surface(&name, &bases);
        }
    }

    pub(super) fn predeclare_class_surface(&mut self, name: &str, parents: &[String]) {
        // Derived from the class NORMALIZED in the same pass — not a second
        // hand-walk over `ClassMember`. There were already three of those
        // (`predeclare_struct_surface`, `compile_normal_class`, each language's
        // normalizer); adding a fourth to fix an ordering bug would have been
        // the same shortcut this work removes.
        let Some(nc) = self.normalized_classes.get(name).cloned() else {
            return;
        };

        // Methods only. Properties are NOT registered: a getter named after a
        // profile value-method (`isEmpty`, `length`, `charAt`) would shadow
        // into the user-method path — measured at 6 dart failures. Fields are
        // NOT registered either: their storage names depend on collision
        // resolution done while the body compiles, and `instance_field_types`
        // feeds `infer_expr_type_hint`, so partial data is worse than none —
        // measured at 2 flutter failures. `compile_normal_class` fills both
        // accurately at definition time.
        let instance_member_names: Vec<String> = nc
            .instance_methods
            .iter()
            .map(|m| self.js_member_storage_name_for_class(name, &m.source_name))
            .collect();
        let static_method_names: Vec<String> = nc
            .static_methods
            .iter()
            .map(|m| self.js_member_storage_name_for_class(name, &m.source_name))
            .collect();

        // A method's SIGNATURE, in the declaration pass — not only when the
        // class body compiles.
        //
        // `compile_normal_class` registers this, but a function body compiles
        // BEFORE any class body does, so `obj.f(y = 9)` inside a `fun` found no
        // signature and every named argument bound positionally. A free
        // function was fine (its own declaration pass ran first), which is what
        // made this look like a class bug rather than an ordering one.
        //
        // Same `CallSignature::from_params` and the same key as the definition
        // site, so the later registration is a no-op repeat rather than a
        // second, different answer. Skipped when an entry already exists: two
        // classes can declare the same method name, and the definition pass
        // still appends the genuine overloads.
        for (m, bound_name) in nc.instance_methods.iter().zip(instance_member_names.iter()) {
            if self.function_signatures.contains_key(bound_name) {
                continue;
            }
            self.function_signatures.insert(
                bound_name.clone(),
                vec![CallSignature::from_params(&m.params)],
            );
        }
        for (m, bound_name) in nc.static_methods.iter().zip(static_method_names.iter()) {
            if self.function_signatures.contains_key(bound_name) {
                continue;
            }
            self.function_signatures.insert(
                bound_name.clone(),
                vec![CallSignature::from_params(&m.params)],
            );
        }

        let bases: Vec<String> = parents.iter().map(|p| self.canon(p)).collect();
        let parent = bases.first().cloned();
        self.pending_classes
            .entry(name.to_string())
            .or_insert(PendingClass {
                parent,
                bases,
                enclosing_class: self.current_class.clone(),
                fields: Vec::new(),
                field_storage_names: HashMap::new(),
                is_value_type: false,
                instance_member_names,
                instance_pointer_method_names: Vec::new(),
                instance_field_types: HashMap::new(),
                static_fields: Vec::new(),
                static_field_types: HashMap::new(),
                static_method_names,
                instance_method_overloads: HashMap::new(),
                static_method_overloads: HashMap::new(),
                nested_types: Vec::new(),
                statics: Vec::new() });
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

        self.defined_globals
            .insert(crate::primitives::classes::ctor_global_for(&name, 0));
        self.pending_classes
            .entry(name.to_string())
            .or_insert(PendingClass {
                parent: None,
                bases: Vec::new(),
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
                statics: Vec::new() });
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
                        _ => None } {
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
                bases: Vec::new(),
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
                statics: Vec::new() },
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
                    name } => {
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
                    target_name } => {
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
                    module_root } => {
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
                    alias: Some(alias) } => {
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
                    } else if self.profile.uses_common_resolver {
                        super::resolver::register_platform_trees();
                        for n in names {
                            let raw_local = n.alias.as_ref().unwrap_or(&n.name).clone();
                            let key = self.canon(&raw_local);
                            let mut segments: Vec<&str> = path
                                .split(['.', '\\'])
                                .filter(|part| !part.is_empty())
                                .collect();
                            segments.push(n.name.as_str());
                            if let Some(target) =
                                crate::primitives::namespaces::resolve_path(&segments)
                            {
                                self.bind_namespace_import_target(key, target);
                                continue;
                            }
                            match self.resolve_namespace_path(&segments) {
                                Some(super::resolver::Resolution::HostImport { module, func })
                                | Some(super::resolver::Resolution::Tree(
                                    crate::primitives::namespaces::ResolutionTarget::HostCall {
                                        module,
                                        func,
                                        ..
                                    },
                                )) => {
                                    self.host_import_bindings.insert(key, (module, func));
                                }
                                Some(super::resolver::Resolution::HostConst(value))
                                | Some(super::resolver::Resolution::Tree(
                                    crate::primitives::namespaces::ResolutionTarget::Const(value),
                                )) => {
                                    self.host_const_bindings.insert(key, value);
                                }
                                Some(super::resolver::Resolution::Tree(target)) => {
                                    self.bind_namespace_import_target(key, target);
                                }
                                _ => {}
                            }
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
                // Simple namespace imports (`Imports System.Text`,
                // `using Flutter.Material`) make bare qualified chains
                // resolve under the mounted namespace tree. The mount itself
                // is profile data (`System` -> `dotnet.system`,
                // `Flutter` -> `flutter`, ...), so this is not a .NET mode.
                crate::ast::ImportKind::Simple { path, alias: None }
                    if self.profile.uses_namespace_resolver() =>
                {
                    if let Some(root) = mounted_ambient_tree_root(&self.tree_mounts, path) {
                        if !self.ambient_tree_roots.iter().any(|p| p == &root) {
                            self.ambient_tree_roots.push(root);
                        }
                    } else if self.profile.namespaces.source_imports_are_namespaces {
                        let source_root = self.canon(&path.replace('\\', "."));
                        if !source_root.is_empty()
                            && !self
                                .source_namespace_imports
                                .iter()
                                .any(|p| p == &source_root)
                        {
                            self.source_namespace_imports.push(source_root);
                        }
                    }
                }
                // Default + other Simple imports: no meaning for host modules; skip.
                crate::ast::ImportKind::Default { .. } | crate::ast::ImportKind::Simple { .. } => {}
            }
        }

        // Redundant name check removed: the loop below only fires on a
        // `__go_named_type(...)` call, which only the Go walker emits.
        {
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
                    _ => None };
                if let Some(type_name) = type_name {
                    self.source_type_aliases.insert(self.canon(name), type_name);
                }
            }
        }
    }

    pub(crate) fn resolve_source_type_alias(&self, name: &str) -> String {
        let normalized = Self::strip_global_namespace_prefix(name);
        let trimmed =
            crate::primitives::generics::erased_type_name(normalized.trim()).replace('\\', ".");
        let (head, tail) = trimmed
            .split_once('.')
            .map(|(head, tail)| (head.trim(), Some(tail.trim())))
            .unwrap_or((trimmed.as_str(), None));
        let (alias_head, suffix) = head
            .strip_suffix("()")
            .map(|bare| (bare.trim_end(), "()"))
            .unwrap_or((head, ""));
        let key = self.canon(alias_head);
        if let Some(target) = self.source_type_aliases.get(&key) {
            return match tail {
                Some(tail) if !tail.is_empty() => format!("{}{}.{}", target, suffix, tail),
                _ => format!("{}{}", target, suffix) };
        }

        let lookup_name = match tail {
            Some(tail) if !tail.is_empty() => format!("{alias_head}.{tail}"),
            _ => alias_head.to_string() };
        if let Some(target) = self.resolve_source_namespace_type(&lookup_name) {
            return format!("{target}{suffix}");
        }

        trimmed
    }

    pub(crate) fn resolve_source_namespace_type(&self, name: &str) -> Option<String> {
        let name = self.canon(&Self::strip_global_namespace_prefix(name).replace('\\', "."));
        if name.is_empty() {
            return None;
        }
        if self.defined_classes.contains(&name) {
            return Some(name);
        }
        for ns in self.source_namespace_contexts() {
            let qualified = self.canon(&format!("{ns}.{name}"));
            if self.defined_classes.contains(&qualified) {
                return Some(qualified);
            }
        }
        for prefix in &self.source_namespace_imports {
            let qualified = self.canon(&format!("{prefix}.{name}"));
            if self.defined_classes.contains(&qualified) {
                return Some(qualified);
            }
        }
        let suffix = format!(".{name}");
        let mut matches = self
            .defined_classes
            .iter()
            .filter(|class_name| class_name.ends_with(&suffix));
        if let Some(qualified) = matches.next().cloned() {
            if matches.next().is_none() {
                return Some(qualified);
            }
        }
        None
    }

    pub(crate) fn resolve_source_namespace_value(&self, name: &str) -> Option<String> {
        let name = self.canon(&Self::strip_global_namespace_prefix(name).replace('\\', "."));
        if name.is_empty() {
            return None;
        }
        if self.defined_functions.contains(&name) || self.defined_globals.contains(&name) {
            return Some(name);
        }
        for ns in self.source_namespace_contexts() {
            let qualified = self.canon(&format!("{ns}.{name}"));
            if self.defined_functions.contains(&qualified)
                || self.defined_globals.contains(&qualified)
            {
                return Some(qualified);
            }
        }
        for prefix in &self.source_namespace_imports {
            let qualified = self.canon(&format!("{prefix}.{name}"));
            if self.defined_functions.contains(&qualified)
                || self.defined_globals.contains(&qualified)
            {
                return Some(qualified);
            }
        }
        None
    }

    fn source_namespace_contexts(&self) -> Vec<String> {
        let mut namespaces = Vec::new();
        if let Some(ns) = self.current_namespace.as_deref() {
            if !ns.is_empty() {
                namespaces.push(self.canon(ns));
            }
        }
        if let Some(class_name) = self.current_class.as_deref() {
            let class_name = self.canon(class_name);
            if let Some((namespace, _)) = class_name.rsplit_once('.') {
                if !namespace.is_empty()
                    && !namespaces
                        .iter()
                        .any(|existing| existing.eq_ignore_ascii_case(namespace))
                {
                    namespaces.push(namespace.to_string());
                }
            }
        }
        namespaces
    }
}
