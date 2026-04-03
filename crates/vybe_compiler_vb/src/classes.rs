use std::rc::Rc;
use vybe_bytecode::{Value, Op};
use vybe_parser_basic::ast::*;
use vybe_compiler_common as common;
use vybe_compiler_common::functions as common_fn;

use crate::compiler::Compiler;
use crate::scope::Scope;

impl Compiler {
    /// Compile a VB Class declaration.
    ///
    /// Strategy (same as JS compiler):
    /// - Class compiles to a constructor function
    /// - Constructor takes `this` as first param, initializes fields, attaches methods
    /// - Property Get/Set → `__get_name` / `__set_name` on this
    /// - Shared members → attached to the constructor function object itself
    /// - Inherits → parent constructor called first to attach parent methods
    /// - `New ClassName(args)` → struct_new + call constructor
    pub(crate) fn compile_class(&mut self, class: &ClassDecl) -> Result<(), String> {
        let name = class.name.as_str();

        // Collect constructor, instance methods, shared methods
        let mut ctor_params: Vec<&Parameter> = Vec::new();
        let mut ctor_body: Vec<&Statement> = Vec::new();
        let mut instance_methods: Vec<&MethodDecl> = Vec::new();
        let mut shared_methods: Vec<&MethodDecl> = Vec::new();

        for method in &class.methods {
            let (is_ctor, is_shared) = match method {
                MethodDecl::Sub(sub) => (sub.name.as_str().eq_ignore_ascii_case("New"), sub.is_shared),
                MethodDecl::Function(func) => (false, func.is_shared),
            };
            if is_ctor {
                if let MethodDecl::Sub(sub) = method {
                    ctor_params = sub.parameters.iter().collect();
                    ctor_body = sub.body.iter().collect();
                }
            } else if is_shared {
                shared_methods.push(method);
            } else {
                instance_methods.push(method);
            }
        }

        // Track parent name for MyBase.New() calls
        let saved_parent = self.current_class_parent.take();
        if let Some(ref parent_type) = class.inherits {
            if let VBType::Custom(parent_name) = parent_type {
                self.current_class_parent = Some(parent_name.to_lowercase());
            }
        }

        // --- Compile the constructor chunk ---
        // Constructor creates its own object — arity is user params only (no Me).
        // This makes cross-language `new X()` work uniformly.
        let chunk = common_fn::create_function_chunk(name, ctor_params.len() as u8);
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        // Params first (VM places them in slots 1..N), then "me" as extra local
        for param in &ctor_params {
            scope.define_local(&param.name.as_str().to_lowercase());
        }
        scope.define_local("me"); // slot after all params

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        let this_slot = self.current_scope().resolve_local("me").unwrap();

        // Track class fields and methods so unresolved names inside methods resolve to Me.field/Me.method
        let saved_fields = std::mem::take(&mut self.class_fields);
        let saved_methods = std::mem::take(&mut self.class_methods);
        for field in &class.fields {
            self.class_fields.insert(field.name.as_str().to_lowercase());
        }
        // Add parent class's fields and methods so derived class can resolve inherited names
        if let Some(ref parent_type) = class.inherits {
            let parent_name = match parent_type {
                VBType::Custom(name) => name.to_lowercase(),
                _ => String::new(),
            };
            if let Some(parent_fields) = self.class_field_map.get(&parent_name) {
                for f in parent_fields {
                    self.class_fields.insert(f.clone());
                }
            }
            if let Some(parent_methods) = self.class_method_map.get(&parent_name) {
                for m in parent_methods {
                    self.class_methods.insert(m.clone());
                }
            }
        }
        for method in &class.methods {
            let method_name = match method {
                MethodDecl::Sub(s) => s.name.as_str().to_lowercase(),
                MethodDecl::Function(f) => f.name.as_str().to_lowercase(),
            };
            if method_name != "new" {
                self.class_methods.insert(method_name);
            }
        }

        // Determine if this class has a real (non-framework) base class
        let has_explicit_ctor = class.methods.iter().any(|m| matches!(m, MethodDecl::Sub(s) if s.name.as_str().eq_ignore_ascii_case("New")));
        let has_user_base = if let Some(ref parent_type) = class.inherits {
            let parent_name = match parent_type {
                VBType::Custom(n) => n.to_lowercase(),
                _ => String::new(),
            };
            let is_framework = parent_name.starts_with("system.")
                || parent_name.contains("windows.forms");
            !parent_name.is_empty() && !is_framework
        } else {
            false
        };

        if has_user_base {
            // ── Child class: call parent constructor (creates the object) ──
            let parent_name = match class.inherits.as_ref().unwrap() {
                VBType::Custom(n) => n.to_lowercase(),
                _ => String::new(),
            };
            if !has_explicit_ctor {
                let parent_idx = self.add_string_constant(&parent_name);
                self.emit_u16(Op::global_get, parent_idx);
                self.emit_u8(Op::call, 0);
                // Store returned object as me
                self.emit_u16(Op::local_set, this_slot);
                self.emit(Op::drop);
            }
            // If explicit ctor: MyBase.New() in the body will handle it
        } else {
            // ── Base class (or framework-derived): create object here ──
            let line = self.line;
            common::classes::emit_new_typed_object(
                &mut self.chunks[self.current_chunk_idx],
                this_slot, name, line,
            );
        }

        // Initialize fields
        for field in &class.fields {
            self.emit_u16(Op::local_get, this_slot);
            if let Some(ref init) = field.initializer {
                self.compile_expression(init)?;
            } else {
                self.emit(Op::null);
            }
            let prop_idx = self.add_string_constant(&field.name.as_str().to_lowercase());
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        // Save parent methods as __base_name before attaching child overrides
        if class.inherits.is_some() {
            for method in &instance_methods {
                let method_name = match method {
                    MethodDecl::Sub(sub) => sub.name.as_str().to_lowercase(),
                    MethodDecl::Function(func) => func.name.as_str().to_lowercase(),
                };
                let line = self.line;
                common::classes::emit_save_base_method(
                    &mut self.chunks[self.current_chunk_idx],
                    this_slot, &method_name, line,
                );
            }
        }

        // Compile instance methods and collect chunk indices for vtable.
        // Methods are BOTH:
        // 1. Attached to this (backward compat — existing code uses struct_get)
        // 2. Registered in the type table (vtable — new GC path)
        let mut method_entries: Vec<(String, usize)> = Vec::new();
        for method in &instance_methods {
            let method_name = match method {
                MethodDecl::Sub(sub) => sub.name.as_str().to_lowercase(),
                MethodDecl::Function(func) => func.name.as_str().to_lowercase(),
            };
            // Compile method chunk — pushes closure ref onto stack
            self.emit_u16(Op::local_get, this_slot);
            self.compile_method_decl(method)?;
            // Record chunk index for type table (the chunk was just added)
            let chunk_idx = self.chunks.len() - 1;
            method_entries.push((method_name.clone(), chunk_idx));
            // Attach to instance (backward compat)
            let prop_idx = self.add_string_constant(&method_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
            // Emit cross-language aliases (e.g. VB tostring → JS toString, Python __str__)
            // Uses 0-upvalue ref_func — safe because aliases point to the same chunk
            // and VB instance methods don't capture constructor-scope upvalues.
            // Skip __get_/__set_ aliases — VB uses explicit method calls, and getter
            // aliases would cause the VM to invoke toString as a property getter,
            // breaking `p.ToString()` (which expects a callable, not a string).
            let line = self.line;
            let aliases = common::classes::cross_language_aliases(&method_name);
            for alias in aliases {
                if *alias != method_name && !alias.starts_with("__get_") && !alias.starts_with("__set_") {
                    common::classes::emit_bind_method(
                        &mut self.chunks[self.current_chunk_idx],
                        this_slot, alias, chunk_idx, line,
                    );
                }
            }
        }

        // Compile constructor body (may call InitializeComponent, MyBase.New, etc.)
        for stmt in &ctor_body {
            self.compile_statement(stmt)?;
        }

        // Wire Handles clauses → emit AddHandler calls
        // e.g. "Sub btn1_Click(...) Handles btn1.Click" → AddHandler(btn1.Click, btn1_click)
        for method in &instance_methods {
            let (method_name, handles) = match method {
                MethodDecl::Sub(sub) => (sub.name.as_str().to_lowercase(), &sub.handles),
                MethodDecl::Function(_) => continue,
            };
            if let Some(handle_list) = handles {
                for handle in handle_list {
                    // handle is like "btn1.Click" or "Me.Load"
                    let parts: Vec<&str> = handle.splitn(2, '.').collect();
                    if parts.len() == 2 {
                        let ctrl = parts[0];
                        let event = parts[1];
                        // Emit: AddHandler(ctrl_name, event_name, method_ref)
                        // Push control name + event name as strings
                        let ctrl_str = if ctrl.eq_ignore_ascii_case("Me") {
                            name.to_lowercase()
                        } else {
                            ctrl.to_lowercase()
                        };
                        self.emit_constant(Value::String(Rc::from(ctrl_str.as_str())));
                        self.emit_constant(Value::String(Rc::from(event)));
                        // Push the method as a closure reference
                        self.emit_u16(Op::local_get, this_slot);
                        let method_idx = self.add_string_constant(&method_name);
                        self.emit_u16(Op::struct_get, method_idx);
                        // Call host onEvent(ctrl, event, handler) — same as AddHandler
                        let import_idx = self.import("vybe:gui", "onEvent");
                        self.emit_host_call(import_idx, 3);
                        self.emit(Op::drop);
                    }
                }
            }
        }

        // Attach property getters/setters
        for prop in &class.properties {
            let prop_name = prop.name.as_str().to_lowercase();

            if let Some(ref getter_body) = prop.getter {
                self.emit_u16(Op::local_get, this_slot);
                self.compile_property_accessor(&format!("get_{}", prop_name), 1, getter_body, true, Some(&prop_name))?;
                let getter_chunk_idx = self.chunks.len() - 1;
                let get_name = format!("__get_{}", prop_name);
                let prop_idx = self.add_string_constant(&get_name);
                self.emit_u16(Op::struct_set, prop_idx);
                self.emit(Op::drop);
                // Emit cross-language aliases for getter (e.g. __get_tostring for Python/JS interop)
                let line = self.line;
                for alias in common::classes::cross_language_aliases(&prop_name) {
                    if *alias != prop_name {
                        let alias_get = format!("__get_{}", alias);
                        common::classes::emit_bind_method(
                            &mut self.chunks[self.current_chunk_idx],
                            this_slot, &alias_get, getter_chunk_idx, line,
                        );
                    }
                }
            }

            if let Some((ref value_param, ref setter_body)) = prop.setter {
                self.emit_u16(Op::local_get, this_slot);
                let param_name = value_param.name.as_str().to_lowercase();
                self.compile_property_accessor_with_param(&format!("set_{}", prop_name), &param_name, setter_body)?;
                let setter_chunk_idx = self.chunks.len() - 1;
                let set_name = format!("__set_{}", prop_name);
                let prop_idx = self.add_string_constant(&set_name);
                self.emit_u16(Op::struct_set, prop_idx);
                self.emit(Op::drop);
                // Emit cross-language aliases for setter
                let line = self.line;
                for alias in common::classes::cross_language_aliases(&prop_name) {
                    if *alias != prop_name {
                        let alias_set = format!("__set_{}", alias);
                        common::classes::emit_bind_method(
                            &mut self.chunks[self.current_chunk_idx],
                            this_slot, &alias_set, setter_chunk_idx, line,
                        );
                    }
                }
            }
        }

        // Stamp/re-stamp type info on this.
        // For base class: emit_new_typed_object already stamped.
        // For child class: re-stamp with child type (parent's stamps are from super).
        if has_user_base {
            let tid_name = format!("__tid_{}", name.to_lowercase());
            let tid_idx = self.add_string_constant(&tid_name);
            self.emit_u16(Op::local_get, this_slot);
            self.emit_u16(Op::global_get, tid_idx);
            self.emit(Op::set_type_id);
            // Update __type string
            self.emit_u16(Op::local_get, this_slot);
            self.emit_constant(Value::String(Rc::from(name)));
            let type_key = self.add_string_constant("__type");
            self.emit_u16(Op::struct_set, type_key);
            self.emit(Op::drop);
        }

        // Return this
        {
            let line = self.line;
            common::classes::emit_constructor_return(
                &mut self.chunks[self.current_chunk_idx],
                this_slot, line,
            );
        }

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.current_class_parent = saved_parent;
        // Store this class's full field/method sets (own + inherited) for future derived classes
        self.class_field_map.insert(name.to_lowercase(), self.class_fields.clone());
        self.class_method_map.insert(name.to_lowercase(), self.class_methods.clone());
        self.class_fields = saved_fields;
        self.class_methods = saved_methods;

        // --- WASM GC: Register type entry in compile-time type table ---
        let parent_name = class.inherits.as_ref().map(|t| match t {
            VBType::Custom(n) => n.to_lowercase(),
            _ => String::new(),
        }).unwrap_or_default();
        let field_names: Vec<String> = class.fields.iter()
            .map(|f| f.name.as_str().to_lowercase())
            .collect();
        let implements: Vec<String> = class.implements.iter().map(|t| match t {
            VBType::Custom(n) => n.to_lowercase(),
            _ => String::new(),
        }).filter(|s| !s.is_empty()).collect();
        let type_entry_idx = self.chunks[0].types.len();
        common::classes::register_type(
            &mut self.chunks,
            name,
            &parent_name,
            field_names,
            method_entries,
            false,
            implements,
            Some(idx),
        );
        self.class_type_ids.insert(name.to_lowercase(), type_entry_idx);

        self.emit_ref_func(idx, &upvalues);

        // If Inherits, copy parent's Shared methods to this constructor
        if let Some(ref parent_type) = class.inherits {
            let parent_name = match parent_type {
                VBType::Custom(pn) => pn.to_lowercase(),
                _ => String::new(),
            };
            if !parent_name.is_empty() && !parent_name.starts_with("system.") {
                self.emit(Op::dup);
                let parent_idx = self.add_string_constant(&parent_name);
                self.emit_u16(Op::global_get, parent_idx);
                let assign_idx = self.import("vybe:object", "assign");
                self.emit_host_call(assign_idx, 2);
                self.emit(Op::drop);
            }
        }

        // --- Attach Shared members to the constructor function itself ---
        for method in &shared_methods {
            let method_name = match method {
                MethodDecl::Sub(sub) => sub.name.as_str().to_lowercase(),
                MethodDecl::Function(func) => func.name.as_str().to_lowercase(),
            };
            self.emit(Op::dup); // keep constructor on stack
            self.compile_shared_method(method)?;
            let prop_idx = self.add_string_constant(&method_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        Ok(())
    }

    /// Compile a method (Sub or Function) as a closure that takes Me as first param.
    fn compile_method_decl(&mut self, method: &MethodDecl) -> Result<(), String> {
        match method {
            MethodDecl::Sub(sub) => {
                let chunk = common_fn::create_function_chunk(sub.name.as_str(), (sub.parameters.len() + 1) as u8);
                let idx = self.chunks.len();
                self.chunks.push(chunk);

                let mut scope = Scope::new_function(); // slot 0 = callee (reserved by new_function)
                scope.define_local("me");              // slot 1 = this (first arg)
                for param in &sub.parameters {
                    scope.define_local(&param.name.as_str().to_lowercase());
                }

                let saved = self.current_chunk_idx;
                self.current_chunk_idx = idx;
                self.scopes.push(scope);

                for stmt in &sub.body { self.compile_statement(stmt)?; }
                common_fn::emit_function_epilogue(&mut self.chunks[idx], self.line);

                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);
            }
            MethodDecl::Function(func) => {
                let chunk = common_fn::create_function_chunk(func.name.as_str(), (func.parameters.len() + 1) as u8);
                let idx = self.chunks.len();
                self.chunks.push(chunk);

                let mut scope = Scope::new_function(); // slot 0 = callee
                scope.define_local("me");              // slot 1 = this
                for param in &func.parameters {
                    scope.define_local(&param.name.as_str().to_lowercase());
                }
                scope.define_local("__return_val");

                let saved = self.current_chunk_idx;
                self.current_chunk_idx = idx;
                self.scopes.push(scope);

                self.function_name_stack.push(func.name.as_str().to_lowercase());
                self.emit(Op::null);
                let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
                self.emit_u16(Op::local_set, rv_slot);
                self.emit(Op::drop);

                for stmt in &func.body { self.compile_statement(stmt)?; }

                self.function_name_stack.pop();
                let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
                self.emit_u16(Op::local_get, rv_slot);
                self.emit(Op::r#return);

                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);
            }
        }
        Ok(())
    }

    /// Compile a Shared method — no Me parameter.
    fn compile_shared_method(&mut self, method: &MethodDecl) -> Result<(), String> {
        match method {
            MethodDecl::Sub(sub) => {
                let chunk = common_fn::create_function_chunk(sub.name.as_str(), sub.parameters.len() as u8);
                let idx = self.chunks.len();
                self.chunks.push(chunk);

                let mut scope = Scope::new_function();
                for param in &sub.parameters {
                    scope.define_local(&param.name.as_str().to_lowercase());
                }

                let saved = self.current_chunk_idx;
                self.current_chunk_idx = idx;
                self.scopes.push(scope);

                for stmt in &sub.body { self.compile_statement(stmt)?; }
                common_fn::emit_function_epilogue(&mut self.chunks[idx], self.line);

                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);
            }
            MethodDecl::Function(func) => {
                let chunk = common_fn::create_function_chunk(func.name.as_str(), func.parameters.len() as u8);
                let idx = self.chunks.len();
                self.chunks.push(chunk);

                let mut scope = Scope::new_function();
                for param in &func.parameters {
                    scope.define_local(&param.name.as_str().to_lowercase());
                }
                scope.define_local("__return_val");

                let saved = self.current_chunk_idx;
                self.current_chunk_idx = idx;
                self.scopes.push(scope);

                self.function_name_stack.push(func.name.as_str().to_lowercase());
                self.emit(Op::null);
                let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
                self.emit_u16(Op::local_set, rv_slot);
                self.emit(Op::drop);

                for stmt in &func.body { self.compile_statement(stmt)?; }

                self.function_name_stack.pop();
                let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
                self.emit_u16(Op::local_get, rv_slot);
                self.emit(Op::r#return);

                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);
            }
        }
        Ok(())
    }

    /// Compile a property getter/setter body as a closure.
    fn compile_property_accessor(&mut self, label: &str, arity: u8, body: &[Statement], has_return: bool, prop_name: Option<&str>) -> Result<(), String> {
        let chunk = common_fn::create_function_chunk(label, arity);
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function(); // slot 0 = callee
        scope.define_local("me");              // slot 1 = this
        if has_return { scope.define_local("__return_val"); }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        if has_return {
            self.emit(Op::null);
            let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
            self.emit_u16(Op::local_set, rv_slot);
            self.emit(Op::drop);
            // Push property name so `PropertyName = value` sets __return_val
            if let Some(pn) = prop_name {
                self.function_name_stack.push(pn.to_string());
            }
        }

        for stmt in body { self.compile_statement(stmt)?; }

        if has_return {
            if prop_name.is_some() {
                self.function_name_stack.pop();
            }
            let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
            self.emit_u16(Op::local_get, rv_slot);
        } else {
            self.emit(Op::null);
        }
        self.emit(Op::r#return);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);
        Ok(())
    }

    /// Compile a property setter (Me + value param).
    fn compile_property_accessor_with_param(&mut self, label: &str, param_name: &str, body: &[Statement]) -> Result<(), String> {
        let chunk = common_fn::create_function_chunk(label, 2); // Me + value
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function(); // slot 0 = callee
        scope.define_local("me");              // slot 1 = this
        scope.define_local(param_name);

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        for stmt in body { self.compile_statement(stmt)?; }
        common_fn::emit_function_epilogue(&mut self.chunks[idx], self.line);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);
        Ok(())
    }
}
