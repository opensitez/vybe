//! Assorted lowering helpers: component-model, type predicates, Fortran ctors, global maps, records, VB statements.
//!
//! Extracted from `compiler/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use super::*;

impl Compiler {
    #[allow(dead_code)]
    pub(super) fn current_offset(&self) -> usize {
        self.chunks[self.current].current_offset()
    }
    pub(crate) fn str_const(&mut self, s: &str) -> u16 {
        self.chunks[self.current].add_constant(Value::String(Arc::from(s)))
    }
    #[allow(dead_code)]
    pub(super) fn shared_str_const(&mut self, s: &str) -> u16 {
        self.chunks[0].add_constant(Value::String(Arc::from(s)))
    }

    pub(crate) fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[self.current].add_import(module, name)
    }
    pub(crate) fn emit_host_call(&mut self, idx: u16, argc: u8) {
        let l = self.line;
        self.chunks[self.current].emit_call(idx, argc, l);
    }

    /// Resolve a qualified identifier to a Component Model host call
    /// `(module, function)` pair when its first segment matches the
    /// profile's `host_packages` list, else `None`.
    ///
    /// Walker conventions: PHP passes backslash-separated names
    /// (`Vybe\Http\Request\method`), other languages should normalize
    /// their separator to `\` before this point (TODO for Python / C# /
    /// etc.). This keeps the resolver language-agnostic.
    ///
    /// Mapping:
    /// - `[Vybe, Http, Request, method]` → `("vybe:http/request", "method")`
    /// - `[Vybe, Math, cos]`             → `("ecma:math", "cos")`
    /// - `[Wasi, Cli, log]`              → `("wasi:logging/logging", "log")`
    ///
    /// First join is `:` (package → interface), further joins use `/`,
    /// last segment is the function name. Everything is lowercased.
    pub(super) fn resolve_component_model_call(&self, name: &str) -> Option<(String, String)> {
        if !name.contains('\\') {
            return None;
        }
        let parts: Vec<&str> = name.split('\\').collect();
        if parts.len() < 2 {
            return None;
        }

        // namespaceplan.md: the global namespace tree is the PRIMARY
        // resolver — it mounts every host export (`vybe.gui.*` next to
        // `ecma.*` and `wasi.*`), so a backslash chain resolves exactly
        // like the dotted chains other languages emit. The manual
        // module-string build below remains the fallback for name shapes
        // the tree doesn't key (e.g. CamelCase segments of kebab-case
        // module names).
        if self.profile.uses_common_resolver {
            match self.resolve_namespace_path(&parts) {
                Some(self::resolver::Resolution::HostImport { module, func }) => {
                    return Some((module, func));
                }
                Some(self::resolver::Resolution::Tree(
                    crate::compiler::namespaces::ResolutionTarget::HostCall { module, func, .. },
                )) => {
                    return Some((module, func));
                }
                _ => {}
            }
        }

        // Consult the Linker's `host_package_roots` map instead of
        // `profile.namespaces.host_packages`. Populated at link time
        // from `EsmDefault::PackageRoot` entries (which the profile
        // loader auto-translates from the legacy list). Component
        // Model package names are lowercase by spec — match
        // case-insensitively regardless of the language's case rules.
        let first_key = parts[0].to_ascii_lowercase();
        if !self.host_package_roots.contains_key(&first_key) {
            return None;
        }

        let lower: Vec<String> = parts.iter().map(|s| s.to_ascii_lowercase()).collect();
        let (func, path) = lower.split_last()?;
        if path.is_empty() {
            return None;
        }

        let module = if path.len() == 1 {
            path[0].clone()
        } else {
            let mut m = path[0].clone();
            m.push(':');
            m.push_str(&path[1]);
            for p in &path[2..] {
                m.push('/');
                m.push_str(p);
            }
            m
        };
        Some((module, func.clone()))
    }

    // ── Crate-private accessors used by `dotnet_register` ──────────────
    //
    // The .NET class registration logic lives in a sibling file
    // (`dotnet_register.rs`) but operates on Compiler internals. These
    // helpers expose just the bits that registration needs without
    // making the underlying fields `pub`.
    pub(crate) fn chunks_mut(&mut self) -> &mut Vec<Chunk> {
        &mut self.chunks
    }
    pub(crate) fn current_line(&self) -> u32 {
        self.line
    }
    pub(crate) fn note_defined_global(&mut self, name: &str) {
        self.defined_globals.insert(name.to_string());
    }
    pub(crate) fn note_defined_class(&mut self, name: &str) {
        self.defined_classes.insert(name.to_string());
    }
    /// Mount a namespace-tree root as ambient (unqualified names resolve under
    /// it) — used when a module imports a platform surface (`flutter.*`).
    pub(crate) fn mount_ambient_root(&mut self, root: &str) {
        if !self.ambient_tree_roots.iter().any(|r| r == root) {
            self.ambient_tree_roots.push(root.to_string());
        }
    }
    pub(crate) fn note_pending_class(&mut self, name: &str, parent: Option<String>) {
        self.pending_classes.insert(
            name.to_string(),
            PendingClass {
                bases: parent.iter().cloned().collect(),
                parent,
                enclosing_class: self.current_class.clone(),
                fields: Vec::new(),
                field_storage_names: HashMap::new(),
                is_value_type: false,
                instance_member_names: Vec::new(),
                instance_pointer_method_names: Vec::new(),
                instance_field_types: HashMap::new(),
                static_fields: Vec::new(),
                static_field_types: HashMap::new(),
                static_method_names: Vec::new(),
                instance_method_overloads: HashMap::new(),
                static_method_overloads: HashMap::new(),
                nested_types: Vec::new(),
                statics: Vec::new(),
            },
        );
    }

    /// Push the canonical event-registry key for a control expression.
    /// Used by AddHandler / RemoveHandler so the GUI host indexes handlers by
    /// the source-stable identifier (field name, class name for `Me`, etc.)
    /// rather than the runtime `.Name` property — renaming a control after
    /// the handler is wired must NOT break dispatch.
    ///
    pub(crate) fn canon(&self, name: &str) -> String {
        let name = if self.is_php_profile() {
            name.strip_prefix('$').unwrap_or(name)
        } else {
            name
        };
        if self.case_sensitive {
            name.to_string()
        } else {
            name.to_lowercase()
        }
    }

    pub(crate) fn php_variable_global_key(&self, name: &str, canon: &str) -> String {
        if self.is_php_profile() && name.starts_with('$') {
            if name.starts_with("$_") {
                name.to_string()
            } else {
                format!("__php_var_{canon}")
            }
        } else {
            canon.to_string()
        }
    }

    /// True if `name` is a class the program actually declares. A real class
    /// of a built-in exception name (e.g. PHP's prelude `LogicException`,
    /// `RuntimeException`, …) must go through the ordinary class emitter, NOT
    /// the `is_exception_type` intrinsic shortcut — otherwise the intrinsic
    /// shape (canonicalized `__type`, no `__types` chain) shadows the real
    /// class and subclass identity is lost.
    pub(crate) fn shadows_builtin_type(&self, name: &str) -> bool {
        self.defined_classes.contains(name)
            || self.defined_classes.contains(&self.canon(name))
            || (!self.case_sensitive
                && self
                    .defined_classes
                    .iter()
                    .any(|g| g.eq_ignore_ascii_case(name)))
    }

    pub(super) fn normalize_type_hint(type_hint: &str) -> String {
        type_hint.trim().to_lowercase()
    }

    pub(super) fn emit_default_value_for_type_hint(&mut self, type_hint: Option<&str>) {
        match type_hint.map(Self::normalize_type_hint).as_deref() {
            Some("integer") | Some("int") | Some("int32") | Some("longint") | Some("real")
            | Some("double") | Some("float") | Some("single") | Some("decimal") | Some("long")
            | Some("int64") | Some("short") | Some("int16") | Some("uint") | Some("uint32")
            | Some("ulong") | Some("uint64") | Some("ushort") | Some("uint16") | Some("byte")
            | Some("sbyte") => inst!(self, core_wasm::f64_const, 0.0),
            Some("char") if self.profile.name == "pascal" => {
                self.emit_const(Value::String(Arc::from("")))
            }
            Some("char") => inst!(self, core_wasm::f64_const, 0.0),
            Some("boolean") | Some("bool") => inst!(self, core_wasm::bool_const, false),
            Some(type_hint) if Self::is_string_type_hint(type_hint) => {
                self.emit_const(Value::String(Arc::from("")))
            }
            _ => self.emit(Op::NULL),
        }
    }

    pub(super) fn is_string_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        normalized == "string"
            || normalized == "system.string"
            || normalized.ends_with(".string")
            || normalized == "character"
            || normalized.starts_with("character(")
            || normalized.starts_with("character*")
    }

    pub(super) fn is_numeric_type_hint(type_hint: &str) -> bool {
        matches!(
            Self::normalize_type_hint(type_hint).as_str(),
            "integer"
                | "int"
                | "int32"
                | "longint"
                | "real"
                | "double"
                | "float"
                | "single"
                | "decimal"
                | "long"
                | "int64"
                | "short"
                | "int16"
                | "uint"
                | "uint32"
                | "ulong"
                | "uint64"
                | "ushort"
                | "uint16"
                | "byte"
                | "sbyte"
        )
    }

    pub(super) fn fortran_out_param_ctor_name(type_hint: &str) -> Option<String> {
        let normalized = Self::normalize_type_hint(type_hint);
        if normalized.ends_with("()")
            || Self::is_numeric_type_hint(&normalized)
            || Self::is_string_type_hint(&normalized)
            || matches!(normalized.as_str(), "boolean" | "bool")
        {
            return None;
        }

        if let Some(inner) = normalized
            .strip_prefix("type(")
            .and_then(|inner| inner.strip_suffix(')'))
        {
            return Some(inner.trim().to_string());
        }

        if let Some(inner) = normalized
            .strip_prefix("class(")
            .and_then(|inner| inner.strip_suffix(')'))
        {
            return Some(inner.trim().to_string());
        }

        Some(normalized)
    }

    pub(super) fn maybe_initialize_fortran_out_param(&mut self, param: &Param) {
        if self.profile.name != "fortran" || param.pass_by != PassBy::Out {
            return;
        }

        let Some(type_hint) = param.type_hint.as_deref() else {
            return;
        };
        let Some(ctor_name) = Self::fortran_out_param_ctor_name(type_hint) else {
            return;
        };
        let Some(slot) = self.scope().resolve(&param.name) else {
            return;
        };
        if !(self.defined_classes.contains(&ctor_name)
            || self.defined_globals.contains(&ctor_name)
            || self.profile.lookup_known_type(&ctor_name).is_some())
        {
            return;
        }

        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);

        if let Some((module, func)) = self
            .profile
            .lookup_known_type(&ctor_name)
            .map(|(module, func)| (module.to_string(), func.to_string()))
        {
            let idx = self.import(&module, &func);
            self.emit_host_call(idx, 0);
        } else {
            let idx = self.global_name_const_idx(&ctor_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u8(Op::CALL_REF, 0);
        }
        self.emit_u16(Op::LOCAL_SET, slot);

        self.chunk().emit_end(line);
    }

    pub(super) fn can_instantiate_fortran_ctor_name(&self, ctor_name: &str) -> bool {
        self.defined_classes.contains(ctor_name)
            || self.defined_globals.contains(ctor_name)
            || self.profile.lookup_known_type(ctor_name).is_some()
    }

    pub(super) fn emit_fortran_ctor_call(&mut self, ctor_name: &str) {
        if let Some((module, func)) = self
            .profile
            .lookup_known_type(ctor_name)
            .map(|(module, func)| (module.to_string(), func.to_string()))
        {
            let idx = self.import(&module, &func);
            self.emit_host_call(idx, 0);
        } else {
            let idx = self.global_name_const_idx(ctor_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u8(Op::CALL_REF, 0);
        }
    }

    pub(super) fn fortran_allocate_ctor_name(&self, target: &Expression) -> Option<String> {
        let type_hint = self.infer_expr_type_hint(target)?;
        let normalized = Self::normalize_type_hint(&type_hint);
        let element_hint = normalized
            .strip_suffix("()")
            .unwrap_or(normalized.as_str())
            .trim();
        let ctor_name = Self::fortran_out_param_ctor_name(element_hint)?;
        self.can_instantiate_fortran_ctor_name(&ctor_name)
            .then_some(ctor_name)
    }

    pub(super) fn emit_fortran_allocated_array(
        &mut self,
        dim_slots: &[u16],
        ctor_name: Option<&str>,
    ) {
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, dim_slots[0]);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__fortran_alloc_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);

        if dim_slots.len() == 1 && ctor_name.is_none() {
            self.emit_u16(Op::LOCAL_GET, array_slot);
            return;
        }

        let idx_slot = self.define_local("__fortran_alloc_idx");
        self.emit_const(Value::F64(0.0));
        self.emit_u16(Op::LOCAL_SET, idx_slot);

        let block_patch = self.chunk().emit_block(line);
        let (loop_patch, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_u16(Op::LOCAL_GET, dim_slots[0]);
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_lt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        self.emit(Op::I32_EQZ);
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        if dim_slots.len() > 1 {
            self.emit_fortran_allocated_array(&dim_slots[1..], ctor_name);
        } else if let Some(ctor_name) = ctor_name {
            self.emit_fortran_ctor_call(ctor_name);
        } else {
            self.emit(Op::NULL);
        }
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_const(Value::F64(1.0));
        self.emit(Op::F64_ADD);
        self.emit_u16(Op::LOCAL_SET, idx_slot);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(loop_patch);
        self.chunk().emit_end(line);
        self.chunk().patch_block(block_patch);

        self.emit_u16(Op::LOCAL_GET, array_slot);
    }

    pub(super) fn expr_prefers_numeric_add(&self, expr: &Expression) -> bool {
        self.infer_expr_type_hint(expr)
            .as_deref()
            .is_some_and(Self::is_numeric_type_hint)
    }

    pub(super) fn compile_expr_with_numeric_add_hint(
        &mut self,
        expr: &Expression,
        prefer_numeric_add: bool,
    ) -> Result<(), String> {
        if prefer_numeric_add {
            if let ExprKind::Binary {
                op: BinOp::Add,
                left,
                right,
            } = &expr.kind
            {
                self.compile_expr_with_numeric_add_hint(left, true)?;
                self.compile_expr_with_numeric_add_hint(right, true)?;
                self.emit(Op::F64_ADD);
                return Ok(());
            }
        }

        self.compile_expr(expr)
    }

    pub(super) fn emit_assignment_type_coercion_for_target(&mut self, target: &Expression) {
        let ExprKind::Ident(name) = &target.kind else {
            return;
        };
        self.emit_assignment_type_coercion_for_ident(name);
    }

    pub(super) fn emit_assignment_type_coercion_for_ident(&mut self, name: &str) {
        if self.lookup_array_binding(name).is_some() {
            return;
        }
        let Some(type_hint) = self.lookup_var_type_hint(name).map(str::to_string) else {
            return;
        };
        let normalized = Self::normalize_type_hint(&type_hint);
        inst!(self, core_wasm::dup);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        let line = self.line;
        self.chunk().emit_if(line);
        match normalized.as_str() {
            "integer" | "int" | "int32" | "longint" | "long" | "int64" | "short" | "int16"
            | "uint" | "uint32" | "ulong" | "uint64" | "ushort" | "uint16" | "byte" | "sbyte" => {
                let number_idx = self.import("ecma:number", "Number");
                self.emit_host_call(number_idx, 1);
                common::convert::emit_to_int(self.chunk(), line);
            }
            "real" | "double" | "float" | "single" | "decimal" => {
                let number_idx = self.import("ecma:number", "Number");
                self.emit_host_call(number_idx, 1);
            }
            _ => {}
        }
        self.chunk().emit_end(line);
    }

    pub(super) fn emit_file_key_compare(&mut self, relation: FileKeyRelation) {
        match relation {
            FileKeyRelation::Equal => {
                let line = self.line;
                crate::compiler::ops::emit_dyn_eq(self.chunk(), line);
            }
            FileKeyRelation::Greater => {
                let line = self.line;
                crate::compiler::ops::emit_dyn_gt(self.chunk(), line);
            }
            FileKeyRelation::GreaterOrEqual => {
                let line = self.line;
                crate::compiler::ops::emit_dyn_ge(self.chunk(), line);
            }
            FileKeyRelation::Less => {
                let line = self.line;
                crate::compiler::ops::emit_dyn_lt(self.chunk(), line);
            }
            FileKeyRelation::LessOrEqual => {
                let line = self.line;
                crate::compiler::ops::emit_dyn_le(self.chunk(), line);
            }
        }
    }

    pub(super) fn emit_global_map_get_into_local(
        &mut self,
        map_name: &str,
        key_slot: u16,
        value_slot: u16,
    ) {
        let map_key = self.shared_global_slot(map_name);
        self.emit_ensure_global_map(map_name);
        self.emit_u16(Op::GLOBAL_GET, map_key);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit(Op::ARRAY_GET);
        self.emit_u16(Op::LOCAL_SET, value_slot);
    }

    pub(super) fn emit_global_map_set_from_local(
        &mut self,
        map_name: &str,
        key_slot: u16,
        value_slot: u16,
    ) {
        let map_key = self.shared_global_slot(map_name);
        self.emit_ensure_global_map(map_name);
        self.emit_u16(Op::GLOBAL_GET, map_key);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::ARRAY_SET);
        self.emit(Op::DROP);
    }

    pub(super) fn emit_global_map_set_const(
        &mut self,
        map_name: &str,
        key_slot: u16,
        value: Value,
    ) {
        let map_key = self.shared_global_slot(map_name);
        self.emit_ensure_global_map(map_name);
        self.emit_u16(Op::GLOBAL_GET, map_key);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_const(value);
        self.emit(Op::ARRAY_SET);
        self.emit(Op::DROP);
    }

    pub(super) fn emit_global_map_set_null(&mut self, map_name: &str, key_slot: u16) {
        let map_key = self.shared_global_slot(map_name);
        self.emit_ensure_global_map(map_name);
        self.emit_u16(Op::GLOBAL_GET, map_key);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit(Op::NULL);
        self.emit(Op::ARRAY_SET);
        self.emit(Op::DROP);
    }

    pub(super) fn emit_record_rows_cache(&mut self, file_slot: u16, rows_slot: u16, len_slot: u16) {
        let line = self.line;
        let path_map_key = self.shared_global_slot("__vb_file_path_by_handle");

        self.emit_global_map_get_into_local("__vb_record_rows_by_handle", file_slot, rows_slot);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        self.emit(Op::REF_IS_NULL);
        self.chunk().emit_if(line);

        self.emit_ensure_global_map("__vb_file_path_by_handle");
        self.emit_u16(Op::GLOBAL_GET, path_map_key);
        self.emit_u16(Op::LOCAL_GET, file_slot);
        self.emit(Op::ARRAY_GET);
        let read_file_idx = self.import("wasi:filesystem", "readFile");
        self.emit_host_call(read_file_idx, 1);
        self.emit_const(Value::String(Arc::from("\n")));
        fn_call!(self, "ecma:string", "split", 2);
        self.emit_u16(Op::LOCAL_SET, rows_slot);

        let skip_trim = self.chunk().emit_block(line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_array_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        inst!(self, core_wasm::i32_const, 0);
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_gt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_not(self.chunk(), line);
        };
        self.chunk().emit_br_if(0, line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        inst!(self, core_wasm::i32_const, 1);
        self.emit(Op::I32_SUB);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        self.emit_const(Value::String(Arc::from("")));
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_eq(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_not(self.chunk(), line);
        };
        self.chunk().emit_br_if(0, line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_pop(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);
        self.chunk().patch_block(skip_trim);

        self.emit_global_map_set_from_local("__vb_record_rows_by_handle", file_slot, rows_slot);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_array_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, len_slot);
    }

    pub(super) fn emit_record_assign_nulls(&mut self, variables: &[String]) {
        for variable in variables {
            self.emit(Op::NULL);
            self.emit_var_set(variable);
        }
    }

    pub(super) fn emit_record_assign_values_from_local(
        &mut self,
        values_slot: u16,
        variables: &[String],
    ) {
        for (index, variable) in variables.iter().enumerate() {
            self.emit_u16(Op::LOCAL_GET, values_slot);
            self.emit_const(Value::F64(index as f64));
            self.emit(Op::ARRAY_GET);
            self.emit_assignment_type_coercion_for_ident(variable);
            self.emit_var_set(variable);
        }
    }

    pub(super) fn emit_record_rewrite_field_format(
        &mut self,
        field_format: Option<&RecordFieldFormat>,
    ) {
        let Some(field_format) = field_format else {
            return;
        };

        let number_idx = self.import("ecma:number", "Number");
        let to_fixed_idx = self.import("ecma:number", "toFixed");
        self.emit_host_call(number_idx, 1);
        self.emit_const(Value::F64(field_format.decimal_places as f64));
        self.emit_host_call(to_fixed_idx, 2);
    }

    pub(super) fn vb_fixed_string_len(type_hint: &str) -> Option<i32> {
        let normalized = Self::normalize_type_hint(type_hint);
        let (base, len) = normalized.split_once('*')?;
        let base = base.trim();
        if base != "string" && base != "system.string" && !base.ends_with(".string") {
            return None;
        }
        len.trim().parse::<i32>().ok().filter(|len| *len >= 0)
    }

    pub(super) fn emit_vb_fixed_string_adjust_from_stack(
        &mut self,
        target_len: i32,
        align_right: bool,
    ) {
        let line = self.line;
        let value_slot = self.define_local("__vb_fixed_string_value");
        let to_string = self.import("ecma:string", "String");
        let pad_idx = self.import(
            "ecma:string",
            if align_right { "padStart" } else { "padEnd" },
        );

        self.emit_host_call(to_string, 1);
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        common::strings::emit_length(self.chunk(), line);
        self.emit_const(Value::I32(target_len));
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_gt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_to_bool(self.chunk(), line);
        }
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        inst!(self, core_wasm::i32_const, 0);
        self.emit_const(Value::I32(target_len));
        common::strings::emit_substring(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_const(Value::I32(target_len));
        self.emit_const(Value::String(Arc::from(" ")));
        self.emit_host_call(pad_idx, 3);
    }

    pub(super) fn compile_vb_fixed_string_stmt(
        &mut self,
        target: &Expression,
        value: &Expression,
        align_right: bool,
    ) -> Result<(), String> {
        let ExprKind::Ident(name) = &target.kind else {
            self.compile_expr(value)?;
            self.emit(Op::DROP);
            return Ok(());
        };
        let Some(type_hint) = self.lookup_var_type_hint(name) else {
            self.compile_expr(value)?;
            self.compile_assign_target(target)?;
            return Ok(());
        };
        let Some(target_len) = Self::vb_fixed_string_len(type_hint) else {
            self.compile_expr(value)?;
            self.compile_assign_target(target)?;
            return Ok(());
        };

        self.compile_expr(value)?;
        self.emit_vb_fixed_string_adjust_from_stack(target_len, align_right);
        self.compile_assign_target(target)
    }

    pub(super) fn compile_vb_mid_stmt(
        &mut self,
        target: &Expression,
        start: &Expression,
        count: &Expression,
        value: &Expression,
    ) -> Result<(), String> {
        let line = self.line;
        let target_slot = self.define_local("__vb_mid_target");
        let start_slot = self.define_local("__vb_mid_start");
        let count_slot = self.define_local("__vb_mid_count");
        let value_slot = self.define_local("__vb_mid_value");
        let prefix_slot = self.define_local("__vb_mid_prefix");
        let replace_slot = self.define_local("__vb_mid_replace");
        let to_string = self.import("ecma:string", "String");

        self.compile_expr(target)?;
        self.emit_u16(Op::LOCAL_SET, target_slot);
        self.compile_expr(start)?;
        common::convert::emit_to_int(self.chunk(), line);
        self.emit_const(Value::I32(1));
        self.emit(Op::I32_SUB);
        self.emit_u16(Op::LOCAL_SET, start_slot);

        self.emit_u16(Op::LOCAL_GET, start_slot);
        inst!(self, core_wasm::i32_const, 0);
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_ge(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_to_bool(self.chunk(), line);
        }
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);
        inst!(self, core_wasm::i32_const, 0);
        self.emit_u16(Op::LOCAL_SET, start_slot);
        self.chunk().emit_end(line);

        self.compile_expr(value)?;
        self.emit_host_call(to_string, 1);
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.compile_expr(count)?;
        common::convert::emit_to_int(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, count_slot);

        self.emit_u16(Op::LOCAL_GET, count_slot);
        self.emit_const(Value::I32(0));
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_ge(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_to_bool(self.chunk(), line);
        }
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, count_slot);
        self.emit_u16(Op::LOCAL_SET, replace_slot);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        common::strings::emit_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, replace_slot);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, target_slot);
        inst!(self, core_wasm::i32_const, 0);
        self.emit_u16(Op::LOCAL_GET, start_slot);
        common::strings::emit_substring(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, prefix_slot);

        self.emit_u16(Op::LOCAL_GET, prefix_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::LOCAL_GET, target_slot);
        self.emit_u16(Op::LOCAL_GET, start_slot);
        self.emit_u16(Op::LOCAL_GET, replace_slot);
        self.emit(Op::I32_ADD);
        self.emit_u16(Op::LOCAL_GET, target_slot);
        common::strings::emit_length(self.chunk(), line);
        common::strings::emit_substring(self.chunk(), line);
        {
            let line = self.line;
            crate::compiler::ops::emit_dyn_add(self.chunk(), line);
        };

        if let ExprKind::Ident(name) = &target.kind {
            if let Some(type_hint) = self.lookup_var_type_hint(name) {
                if let Some(target_len) = Self::vb_fixed_string_len(type_hint) {
                    self.emit_vb_fixed_string_adjust_from_stack(target_len, false);
                }
            }
        }

        self.compile_assign_target(target)
    }

    pub(super) fn compile_vb_err_raise_stmt(&mut self, args: &[Argument]) -> Result<(), String> {
        if let Some(description) = args.get(2).or_else(|| args.get(1)).or_else(|| args.first()) {
            self.compile_expr(&description.value)?;
        } else {
            self.emit_const(Value::String(Arc::from("")));
        }

        self.emit_js_exception_ctor_from_message_value("Exception")?;

        if let Some(number) = args.first() {
            inst!(self, core_wasm::dup);
            self.compile_expr(&number.value)?;
            let key = self.str_const("number");
            self.emit_u16(Op::STRUCT_SET, key);
            self.emit(Op::DROP);
        }

        if let Some(source) = args.get(1) {
            inst!(self, core_wasm::dup);
            self.compile_expr(&source.value)?;
            let key = self.str_const("source");
            self.emit_u16(Op::STRUCT_SET, key);
            self.emit(Op::DROP);
        }

        let line = self.line;
        common::errors::emit_throw(self.chunk(), line);
        Ok(())
    }

    pub(super) fn is_collection_like_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        let bare = normalized
            .split('<')
            .next()
            .unwrap_or(normalized.as_str())
            .trim_end_matches('?');
        let terminal = bare.rsplit('.').next().unwrap_or(bare);
        Self::is_string_type_hint(type_hint)
            || matches!(
                terminal,
                "list"
                    | "arraylist"
                    | "dictionary"
                    | "queue"
                    | "stack"
                    | "hashset"
                    | "sortedset"
                    | "set"
                    | "collection"
                    | "icollection"
                    | "readonlycollection"
                    | "enumerable"
                    | "ienumerable"
                    | "readonlylist"
                    | "ilist"
                    | "array"
            )
            || bare.ends_with("[]")
            || normalized.ends_with("()")
    }
}
