use std::rc::Rc;
use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_basic::ast::*;

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

        // --- Compile the constructor chunk ---
        let mut chunk = Chunk::new(name);
        chunk.arity = (1 + ctor_params.len()) as u8; // Me + ctor params
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("me");
        for param in &ctor_params {
            scope.define_local(&param.name.as_str().to_lowercase());
        }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        let this_slot = self.current_scope().resolve_local("me").unwrap();

        // If Inherits, store parent constructor as __super on this
        // The actual call happens via MyBase.New() in the constructor body
        if let Some(ref parent_type) = class.inherits {
            let parent_name = match parent_type {
                VBType::Custom(name) => name.to_lowercase(),
                _ => String::new(),
            };
            if !parent_name.is_empty() {
                // Store __super = parent constructor on this
                self.emit_u16(Op::local_get, this_slot);
                let parent_idx = self.add_string_constant(&parent_name);
                self.emit_u16(Op::global_get, parent_idx);
                let super_idx = self.add_string_constant("__super");
                self.emit_u16(Op::struct_set, super_idx);
                self.emit(Op::drop);
            }
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

        // Compile constructor body (may contain MyBase.New() calls)
        for stmt in &ctor_body {
            self.compile_statement(stmt)?;
        }

        // Attach instance methods
        for method in &instance_methods {
            let method_name = match method {
                MethodDecl::Sub(sub) => sub.name.as_str().to_lowercase(),
                MethodDecl::Function(func) => func.name.as_str().to_lowercase(),
            };
            self.emit_u16(Op::local_get, this_slot);
            self.compile_method_decl(method)?;
            let prop_idx = self.add_string_constant(&method_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        // Attach property getters/setters
        for prop in &class.properties {
            let prop_name = prop.name.as_str().to_lowercase();

            if let Some(ref getter_body) = prop.getter {
                self.emit_u16(Op::local_get, this_slot);
                self.compile_property_accessor(&format!("get_{}", prop_name), 1, getter_body, true)?;
                let get_name = format!("__get_{}", prop_name);
                let prop_idx = self.add_string_constant(&get_name);
                self.emit_u16(Op::struct_set, prop_idx);
                self.emit(Op::drop);
            }

            if let Some((ref value_param, ref setter_body)) = prop.setter {
                self.emit_u16(Op::local_get, this_slot);
                let param_name = value_param.name.as_str().to_lowercase();
                self.compile_property_accessor_with_param(&format!("set_{}", prop_name), &param_name, setter_body)?;
                let set_name = format!("__set_{}", prop_name);
                let prop_idx = self.add_string_constant(&set_name);
                self.emit_u16(Op::struct_set, prop_idx);
                self.emit(Op::drop);
            }
        }

        // Return this
        self.emit_u16(Op::local_get, this_slot);
        self.emit(Op::r#return);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);

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
                let mut chunk = Chunk::new(sub.name.as_str());
                chunk.arity = (sub.parameters.len() + 1) as u8;
                let idx = self.chunks.len();
                self.chunks.push(chunk);

                let mut scope = Scope::new_function();
                scope.define_local("me");
                for param in &sub.parameters {
                    scope.define_local(&param.name.as_str().to_lowercase());
                }

                let saved = self.current_chunk_idx;
                self.current_chunk_idx = idx;
                self.scopes.push(scope);

                for stmt in &sub.body { self.compile_statement(stmt)?; }
                self.emit(Op::null);
                self.emit(Op::r#return);

                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);
            }
            MethodDecl::Function(func) => {
                let mut chunk = Chunk::new(func.name.as_str());
                chunk.arity = (func.parameters.len() + 1) as u8;
                let idx = self.chunks.len();
                self.chunks.push(chunk);

                let mut scope = Scope::new_function();
                scope.define_local("me");
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
                let mut chunk = Chunk::new(sub.name.as_str());
                chunk.arity = sub.parameters.len() as u8;
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
                self.emit(Op::null);
                self.emit(Op::r#return);

                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);
            }
            MethodDecl::Function(func) => {
                let mut chunk = Chunk::new(func.name.as_str());
                chunk.arity = func.parameters.len() as u8;
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
    fn compile_property_accessor(&mut self, label: &str, arity: u8, body: &[Statement], has_return: bool) -> Result<(), String> {
        let mut chunk = Chunk::new(label);
        chunk.arity = arity;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("me");
        if has_return { scope.define_local("__return_val"); }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        if has_return {
            self.emit(Op::null);
            let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
            self.emit_u16(Op::local_set, rv_slot);
            self.emit(Op::drop);
        }

        for stmt in body { self.compile_statement(stmt)?; }

        if has_return {
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
        let mut chunk = Chunk::new(label);
        chunk.arity = 2; // Me + value
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("me");
        scope.define_local(param_name);

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        for stmt in body { self.compile_statement(stmt)?; }
        self.emit(Op::null);
        self.emit(Op::r#return);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);
        Ok(())
    }
}
