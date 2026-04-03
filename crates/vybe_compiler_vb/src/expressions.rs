use std::rc::Rc;
use vybe_bytecode::{Value, Op};
use vybe_compiler_common::collections as common_collections;
use vybe_compiler_common::expressions as common_expr;
use vybe_compiler_common::functions as common_fn;
use vybe_parser_basic::ast::*;

use crate::compiler::{Compiler, VarResolution};
use crate::scope::Scope;

impl Compiler {
    pub(crate) fn compile_expression(&mut self, expr: &Expression) -> Result<(), String> {
        match expr {
            Expression::IntegerLiteral(n) => {
                if *n == 0 {
                    self.emit(Op::i32_const_0);
                } else if *n == 1 {
                    self.emit(Op::i32_const_1);
                } else {
                    self.emit_constant(Value::I32(*n));
                }
            }
            Expression::DoubleLiteral(n) => {
                let v = *n;
                if v.fract() == 0.0 && v >= i32::MIN as f64 && v <= i32::MAX as f64 {
                    if v == 0.0 { self.emit(Op::i32_const_0); }
                    else if v == 1.0 { self.emit(Op::i32_const_1); }
                    else { self.emit_constant(Value::I32(v as i32)); }
                } else {
                    self.emit_constant(Value::F64(v));
                }
            }
            Expression::StringLiteral(s) => self.emit_constant(Value::String(Rc::from(s.as_str()))),
            Expression::BooleanLiteral(b) => {
                if *b { self.emit(Op::r#true); } else { self.emit(Op::r#false); }
            }
            Expression::Nothing => self.emit(Op::null),

            // Me (this) reference inside a class method
            Expression::Me => {
                match self.resolve_variable("me") {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Global => self.emit(Op::null),
                }
            }

            // MyBase reference — used for MyBase.Method() calls
            Expression::MyBase => {
                // MyBase resolves to Me — the parent's methods are already on the object
                // For MyBase.New() the compiler handles it specially in method calls
                match self.resolve_variable("me") {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Global => self.emit(Op::null),
                }
            }

            Expression::Variable(id) => {
                let name = id.as_str().to_lowercase();
                match self.resolve_variable(&name) {
                    VarResolution::Local(slot) => {
                        self.emit_u16(Op::local_get, slot);
                        // ByRef params are boxes — dereference with array_get 0
                        if self.current_scope().is_byref(&name) {
                            self.emit(Op::i32_const_0);
                            common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                        }
                    }
                    VarResolution::Global => {
                        // Inside a class: unresolved name that's a field → Me.field
                        if self.class_fields.contains(&name) {
                            if let Some(me_slot) = self.current_scope().resolve_local("me") {
                                self.emit_u16(Op::local_get, me_slot);
                                let prop_idx = self.add_string_constant(&name);
                                self.emit_u16(Op::struct_get, prop_idx);
                            } else {
                                let idx = self.add_string_constant(&name);
                                self.emit_u16(Op::global_get, idx);
                            }
                        } else {
                            let idx = self.add_string_constant(&name);
                            self.emit_u16(Op::global_get, idx);
                        }
                    }
                }
            }
            Expression::MemberAccess(obj, member) => {
                // Fallback: try builtin method table
                if let Expression::Variable(ref obj_name) = **obj {
                    let obj_lower = obj_name.as_str().to_lowercase();
                    let mem_lower = member.as_str().to_lowercase();
                    let full = format!("{}.{}", obj_lower, mem_lower);
                    if let Some(()) = self.try_compile_builtin_method(&full, &[])? {
                        // Compiled as a 0-arg host call
                    } else {
                        self.compile_expression(obj)?;
                        let idx = self.add_string_constant(&mem_lower);
                        self.emit_u16(Op::struct_get, idx);
                    }
                } else {
                    self.compile_expression(obj)?;
                    let idx = self.add_string_constant(&member.as_str().to_lowercase());
                    self.emit_u16(Op::struct_get, idx);
                }
            }
            Expression::ArrayAccess(arr, indices) => {
                let name = arr.as_str().to_lowercase();
                match self.resolve_variable(&name) {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(&name);
                        self.emit_u16(Op::global_get, idx);
                    }
                }
                if let Some(index) = indices.first() {
                    self.compile_expression(index)?;
                    common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                }
            }
            Expression::ArrayLiteral(elems) => {
                for e in elems { self.compile_expression(e)?; }
                common_collections::emit_array_new(&mut self.chunks[self.current_chunk_idx], elems.len() as u16, self.line);
            }

            // Arithmetic
            Expression::Add(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::dyn_add); }
            Expression::Subtract(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::f64_sub); }
            Expression::Multiply(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::f64_mul); }
            Expression::Divide(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::f64_div); }
            Expression::IntegerDivide(a, b) => {
                self.compile_expression(a)?;
                self.emit(Op::i32_from_f64);
                self.compile_expression(b)?;
                self.emit(Op::i32_from_f64);
                self.emit(Op::i32_div_s);
            }
            Expression::Modulo(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::f64_mod); }
            Expression::Exponent(a, b) => {
                self.compile_expression(a)?; self.compile_expression(b)?;
                let idx = self.import("vybe:math", "pow");
                self.emit_host_call(idx, 2);
            }
            Expression::Concatenate(a, b) => {
                self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::str_concat);
            }
            Expression::Negate(a) => { self.compile_expression(a)?; self.emit(Op::dyn_neg); }
            Expression::Not(a) => { self.compile_expression(a)?; self.emit(Op::dyn_not); }

            // Comparison
            Expression::Equal(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::dyn_eq); }
            Expression::NotEqual(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::dyn_ne); }
            Expression::LessThan(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::dyn_lt); }
            Expression::LessThanOrEqual(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::dyn_le); }
            Expression::GreaterThan(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::dyn_gt); }
            Expression::GreaterThanOrEqual(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::dyn_ge); }

            // Logical (short-circuit)
            Expression::And(a, b) | Expression::AndAlso(a, b) => {
                self.compile_expression(a)?;
                let jump = common_expr::emit_and_start(&mut self.chunks[self.current_chunk_idx], self.line);
                self.compile_expression(b)?;
                common_expr::emit_short_circuit_end(&mut self.chunks[self.current_chunk_idx], jump);
            }
            Expression::Or(a, b) | Expression::OrElse(a, b) => {
                self.compile_expression(a)?;
                let jump = common_expr::emit_or_start(&mut self.chunks[self.current_chunk_idx], self.line);
                self.compile_expression(b)?;
                common_expr::emit_short_circuit_end(&mut self.chunks[self.current_chunk_idx], jump);
            }

            // Bitwise operators
            Expression::Xor(a, b) => {
                self.compile_expression(a)?; self.compile_expression(b)?;
                self.emit(Op::i32_xor);
            }
            Expression::BitShiftLeft(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::i32_shl); }
            Expression::BitShiftRight(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::i32_shr_s); }

            // Is / IsNot — reference equality
            Expression::Is(a, b) => {
                self.compile_expression(a)?; self.compile_expression(b)?;
                self.emit(Op::dyn_eq);
            }
            Expression::IsNot(a, b) => {
                self.compile_expression(a)?; self.compile_expression(b)?;
                self.emit(Op::dyn_ne);
            }

            // TypeOf expr Is Type — resolved via TypeRegistry
            Expression::TypeOf { expr, type_name } => {
                self.compile_expression(expr)?;
                // Emit the target type name as a constant, then ref_is_type opcode
                let type_idx = self.add_string_constant(&type_name.to_lowercase());
                self.emit_u16(Op::ref_test, type_idx);
            }

            // Like — string pattern matching (simplified)
            Expression::Like(a, b) => {
                self.compile_expression(a)?; self.compile_expression(b)?;
                // Simplified: treat as string equality
                self.emit(Op::dyn_eq);
            }

            // AddressOf funcName (stored as string by parser)
            Expression::AddressOf(name) => {
                let func_name = name.to_lowercase();
                let idx = self.add_string_constant(&func_name);
                self.emit_u16(Op::global_get, idx);
            }

            // Date literal
            Expression::DateLiteral(s) => {
                self.emit_constant(Value::String(Rc::from(s.as_str())));
            }

            Expression::Await(expr) => {
                self.compile_expression(expr)?;
                common_fn::emit_await(&mut self.chunks[self.current_chunk_idx], self.line);
            }

            // WithTarget — reference to the With block's target object
            Expression::WithTarget => {
                match self.resolve_variable("__with_obj") {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    _ => self.emit(Op::null),
                }
            }

            // Query (LINQ) — not yet compiled, emit null
            Expression::Query(_) => {
                self.emit(Op::null);
            }

            // XML literal — emit as null (would need XML serialization)
            Expression::XmlLiteral(_) => {
                self.emit(Op::null);
            }

            // New with initializers — compile New then set properties
            Expression::NewWithInitializer(class_name, args, inits) => {
                self.compile_new_expr(class_name, args)?;
                for (prop, val) in inits {
                    self.emit(Op::dup);
                    self.compile_expression(val)?;
                    let idx = self.add_string_constant(&prop.to_lowercase());
                    self.emit_u16(Op::struct_set, idx);
                    self.emit(Op::drop);
                }
            }

            // New collection from initializer
            Expression::NewFromInitializer(class_name, args, items) => {
                self.compile_new_expr(class_name, args)?;
                // Push items into the collection — simplified: just create array
                for item in items {
                    self.compile_expression(item)?;
                }
                if !items.is_empty() {
                    common_collections::emit_array_new(&mut self.chunks[self.current_chunk_idx], items.len() as u16, self.line);
                }
            }

            // Function call
            Expression::Call(name, args) => {
                self.compile_call_expr(name, args)?;
            }
            Expression::MethodCall(obj, method, args) => {
                self.compile_method_call(obj, method, args)?;
            }
            Expression::New(class_name, args) => {
                self.compile_new_expr(class_name, args)?;
            }

            // If expression (ternary)
            Expression::IfExpression(cond, then_val, else_val) => {
                self.compile_expression(cond)?;
                let false_jump = common_expr::emit_ternary_start(&mut self.chunks[self.current_chunk_idx], self.line);
                self.compile_expression(then_val)?;
                let end_jump = common_expr::emit_ternary_middle(&mut self.chunks[self.current_chunk_idx], false_jump, self.line);
                if let Some(ev) = else_val {
                    self.compile_expression(ev)?;
                } else {
                    self.emit(Op::null);
                }
                common_expr::emit_ternary_end(&mut self.chunks[self.current_chunk_idx], end_jump);
            }

            // Cast
            Expression::Cast { expr, .. } => {
                self.compile_expression(expr)?;
            }

            // Lambda
            Expression::Lambda { params, body } => {
                self.compile_lambda(params, body)?;
            }

        }
        Ok(())
    }

    fn compile_method_call(&mut self, obj: &Expression, method: &Identifier, args: &[Expression]) -> Result<(), String> {
        // WinForms no-op methods — layout, refresh, etc.
        let m = method.as_str().to_lowercase();
        if matches!(m.as_str(),
            "suspendlayout" | "resumelayout" | "performlayout" |
            "refresh" | "invalidate" | "update" | "begininit" | "endinit" |
            "dispose" | "select" | "focus" | "bringtofront" | "sendtoback"
        ) {
            self.emit(Op::null);
            return Ok(());
        }

        // Component Model: try interface resolution on the full chain
        // e.g. System.Math.Floor(3.7) → MemberAccess(System, Math) + method "Floor"
        //   → flatten to ["System", "Math", "Floor"] → resolve "system.math" interface → "floor"
        {
            let mut parts = Self::flatten_member_chain(obj);
            if !parts.is_empty() {
                parts.push(method.as_str().to_string());
                // Try WASM intrinsic first (e.g. Math.Floor → f64_floor opcode)
                let dotted = parts.iter().map(|s| s.to_lowercase()).collect::<Vec<_>>().join(".");
                if let Some(()) = self.try_compile_wasm_intrinsic(&dotted, args)? {
                    return Ok(());
                }
                let part_refs: Vec<&str> = parts.iter().map(|s| s.as_str()).collect();
                if let Some((module, func)) = self.resolve_interface_call(&part_refs) {
                    // Direct call_import — compile-time resolved, no struct_get chain
                    for arg in args { self.compile_expression(arg)?; }
                    let idx = self.import(&module, &func);
                    self.emit_host_call(idx, args.len() as u8);
                    return Ok(());
                }
                // Namespace object resolution: chain struct_get for each part
                let lower_parts: Vec<String> = parts.iter().map(|s| s.to_lowercase()).collect();
                if self.is_namespace(&lower_parts[0]) && lower_parts[0] != "console" {
                    let root_idx = self.add_string_constant(&lower_parts[0]);
                    self.emit_u16(Op::global_get, root_idx);
                    for part in &lower_parts[1..] {
                        let idx = self.add_string_constant(part);
                        self.emit_u16(Op::struct_get, idx);
                    }
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, args.len() as u8);
                    return Ok(());
                }
            }
        }

        // MyBase.New(args) — call parent constructor (which creates and returns an object).
        // Result replaces Me in the child constructor.
        if matches!(obj, Expression::MyBase) && method.as_str().eq_ignore_ascii_case("New") {
            if let Some(parent) = self.current_class_parent.clone() {
                let parent_idx = self.add_string_constant(&parent);
                self.emit_u16(Op::global_get, parent_idx);
                for arg in args { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, args.len() as u8);
                // Store returned object as Me
                if let Some(me_slot) = self.current_scope().resolve_local("me") {
                    self.emit_u16(Op::local_set, me_slot);
                    self.emit(Op::drop);
                }
            }
            return Ok(());
        }

        // MyBase.Method(args) — call parent's version via __base_method
        if matches!(obj, Expression::MyBase) {
            let meth_lower = method.as_str().to_lowercase();
            let base_name = format!("__base_{}", meth_lower);
            match self.resolve_variable("me") {
                VarResolution::Local(slot) => {
                    self.emit_u16(Op::local_get, slot);
                    let prop_idx = self.add_string_constant(&base_name);
                    self.emit_u16(Op::struct_get, prop_idx);
                    self.emit_u16(Op::local_get, slot);
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, (args.len() + 1) as u8);
                }
                _ => { self.emit(Op::null); }
            }
            return Ok(());
        }

        // Me.Method() — method call on Me object
        if matches!(*obj, Expression::Me) {
            if let Some(me_slot) = self.current_scope().resolve_local("me") {
                let meth_lower = method.as_str().to_lowercase();
                match meth_lower.as_str() {
                    // No-op layout methods
                    "suspendlayout" | "resumelayout" | "performlayout"
                    | "refresh" | "invalidate" | "update" | "begininit" | "endinit"
                    | "dispose" | "select" | "bringtofront" | "sendtoback" => {
                        self.emit(Op::null);
                        return Ok(());
                    }
                    _ => {
                        // Try object method first (controls have show/close/focus on them)
                        // If not found at runtime, falls through gracefully
                        self.emit_u16(Op::local_get, me_slot);
                        let prop_idx = self.add_string_constant(&meth_lower);
                        self.emit_u16(Op::struct_get, prop_idx);
                        self.emit_u16(Op::local_get, me_slot); // this
                        for arg in args { self.compile_expression(arg)?; }
                        self.emit_u8(Op::call, (args.len() + 1) as u8);
                        return Ok(());
                    }
                }
            }
        }

        if let Expression::Variable(ref obj_name) = *obj {
            let obj_lower = obj_name.as_str().to_lowercase();
            let meth_lower = method.as_str().to_lowercase();

            // Me.Show() / Me.Close() inside a class → emit GUI host calls
            if obj_lower == "me" && self.current_scope().resolve_local("me").is_some() {
                let me_slot = self.current_scope().resolve_local("me").unwrap();
                match meth_lower.as_str() {
                    "show" => {
                        self.emit_u16(Op::local_get, me_slot);
                        let idx = self.import("vybe:gui", "showForm");
                        self.emit_host_call(idx, 1);
                        return Ok(());
                    }
                    "close" => {
                        self.emit_u16(Op::local_get, me_slot);
                        let idx = self.import("vybe:gui", "closeForm");
                        self.emit_host_call(idx, 1);
                        return Ok(());
                    }
                    "showdialog" => {
                        self.emit_u16(Op::local_get, me_slot);
                        let idx = self.import("vybe:gui", "showFormDialog");
                        self.emit_host_call(idx, 1);
                        return Ok(());
                    }
                    _ => {}
                }
            }

            let full_name = format!("{}.{}", obj_lower, meth_lower);
            if let Some(result) = self.try_compile_builtin_method(&full_name, args)? {
                let _ = result;
            } else if self.is_namespace(&obj_lower) {
                // Namespace call — no `this`
                self.compile_expression(obj)?;
                let prop_idx = self.add_string_constant(&meth_lower);
                self.emit_u16(Op::struct_get, prop_idx);
                for arg in args { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, args.len() as u8);
            } else if self.defined_classes.contains(&obj_lower)
                && !matches!(self.resolve_variable(&obj_lower), VarResolution::Local(_))
                && !self.class_fields.contains(&obj_lower)
            {
                // Static class call — name is ONLY a class, not a variable or field
                self.compile_expression(obj)?;
                let prop_idx = self.add_string_constant(&meth_lower);
                self.emit_u16(Op::struct_get, prop_idx);
                for arg in args { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, args.len() as u8);
            } else {
                self.compile_expression(obj)?;
                let prop_idx = self.add_string_constant(&meth_lower);
                self.emit_u16(Op::struct_get, prop_idx);
                self.compile_expression(obj)?;
                for arg in args { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, (args.len() + 1) as u8);
            }
        } else {
            // Check for form.Controls.Add(ctrl) pattern
            if let Expression::MemberAccess(parent, member) = obj {
                let member_lower = member.as_str().to_lowercase();
                let meth_lower = method.as_str().to_lowercase();
                if member_lower == "controls" && meth_lower == "add" {
                    // Pass parent object directly — controlsAdd extracts name from it
                    self.compile_expression(parent)?;
                    for arg in args { self.compile_expression(arg)?; }
                    let import_idx = self.import("vybe:gui", "controlsAdd");
                    self.emit_host_call(import_idx, (args.len() + 1) as u8);
                    return Ok(());
                }
            }

            if self.is_namespace_expr(obj) {
                let meth_lower = method.as_str().to_lowercase();
                self.compile_expression(obj)?;
                let prop_idx = self.add_string_constant(&meth_lower);
                self.emit_u16(Op::struct_get, prop_idx);
                for arg in args { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, args.len() as u8);
            } else {
                let meth_lower = method.as_str().to_lowercase();
                self.compile_expression(obj)?;
                let prop_idx = self.add_string_constant(&meth_lower);
                self.emit_u16(Op::struct_get, prop_idx);
                self.compile_expression(obj)?;
                for arg in args { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, (args.len() + 1) as u8);
            }
        }
        Ok(())
    }

    fn compile_new_expr(&mut self, class_name: &Identifier, args: &[Expression]) -> Result<(), String> {
        // Strip trailing "()" if parser included it in the name
        let raw = class_name.as_str().to_lowercase();
        let name = raw.trim_end_matches("()").trim_end_matches('(').to_string();

        // Built-in exception types
        if matches!(name.as_str(), "exception" | "argumentexception" | "invalidoperationexception"
            | "notimplementedexception" | "notsupportedexception") {
            self.emit_u16(Op::struct_new, 0);
            self.emit(Op::dup);
            if let Some(msg_arg) = args.first() {
                self.compile_expression(msg_arg)?;
            } else {
                self.emit_constant(Value::String(Rc::from("")));
            }
            let msg_idx = self.add_string_constant("message");
            self.emit_u16(Op::struct_set, msg_idx);
            self.emit(Op::drop);
            return Ok(());
        }

        // User-defined class takes priority over built-in types.
        // Constructor creates its own object — just call(argc) with no pre-created this.
        if self.defined_classes.contains(&name) {
            let idx = self.add_string_constant(&name);
            self.emit_u16(Op::global_get, idx);
            for arg in args { self.compile_expression(arg)?; }
            self.emit_u8(Op::call, args.len() as u8);
            return Ok(());
        }

        // Strip generic type params and fully-qualified prefixes
        let name = name.find("(of ").map(|p| name[..p].to_string()).unwrap_or(name);
        let bare = name
            .strip_prefix("system.data.sqlclient.").or_else(|| name.strip_prefix("system.data.oledb."))
            .or_else(|| name.strip_prefix("system.net.sockets."))
            .or_else(|| name.strip_prefix("system.io."))
            .or_else(|| name.strip_prefix("system.collections."))
            .or_else(|| name.strip_prefix("system.text."))
            .or_else(|| name.strip_prefix("system.windows.forms."))
            .or_else(|| name.strip_prefix("system.drawing."))
            .or_else(|| name.strip_prefix("adodb."))
            .unwrap_or(&name)
            .to_string();

        // 1. TypeRegistry known_types table — direct call, no null prefix
        if let Some(&(module, func)) = self.known_types.get(&bare) {
            for arg in args { self.compile_expression(arg)?; }
            let idx = self.import(module, func);
            self.emit_host_call(idx, args.len() as u8);
            // Stamp type_id from __tid_<name> global (WASM GC)
            self.emit_set_type_id(&bare);
            return Ok(());
        }

        // 2. WinForms controls
        let capitalized = capitalize_control_name(&bare);
        if !capitalized.is_empty() && capitalized != bare {
            for arg in args { self.compile_expression(arg)?; }
            let hn = format!("new_{}", capitalized);
            let idx = self.import("vybe:gui", &hn);
            self.emit_host_call(idx, args.len() as u8);
            // Stamp type_id from __tid_<controltype> global (WASM GC)
            self.emit_set_type_id(&capitalized.to_lowercase());
            return Ok(());
        }

        // 3. Cross-language class: WASM import
        let idx = self.import("*", &name);
        for arg in args { self.compile_expression(arg)?; }
        self.emit_host_call(idx, args.len() as u8);
        Ok(())
    }

    fn compile_lambda(&mut self, params: &[Parameter], body: &LambdaBody) -> Result<(), String> {
        let chunk = common_fn::create_function_chunk("<lambda>", params.len() as u8);
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        for param in params {
            scope.define_local(&param.name.as_str().to_lowercase());
        }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        match body {
            LambdaBody::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::r#return);
            }
            LambdaBody::Statement(stmt) => {
                self.compile_statement(stmt)?;
                self.emit(Op::null);
                self.emit(Op::r#return);
            }
            LambdaBody::Block(stmts) => {
                for s in stmts { self.compile_statement(s)?; }
                self.emit(Op::null);
                self.emit(Op::r#return);
            }
        }

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;

        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in &upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }
        Ok(())
    }
}

impl Compiler {
    /// Flatten a MemberAccess chain into parts.
    /// `System.Windows.Forms.Button` → ["System", "Windows", "Forms", "Button"]
    pub(crate) fn flatten_member_chain(expr: &Expression) -> Vec<String> {
        match expr {
            Expression::Variable(name) => vec![name.as_str().to_string()],
            Expression::MemberAccess(inner, member) => {
                let mut parts = Self::flatten_member_chain(inner);
                parts.push(member.as_str().to_string());
                parts
            }
            _ => vec![],
        }
    }
}

/// Capitalize control type name: "textbox" → "TextBox", "datagridview" → "DataGridView"
pub fn capitalize_control_name(name: &str) -> String {
    // Map of known control names with proper casing
    match name {
        "button" => "Button", "label" => "Label", "textbox" => "TextBox",
        "checkbox" => "CheckBox", "radiobutton" => "RadioButton",
        "combobox" => "ComboBox", "listbox" => "ListBox",
        "panel" => "Panel", "groupbox" => "GroupBox",
        "tabcontrol" => "TabControl", "tabpage" => "TabPage",
        "datagridview" => "DataGridView", "progressbar" => "ProgressBar",
        "trackbar" => "TrackBar", "numericupdown" => "NumericUpDown",
        "datetimepicker" => "DateTimePicker", "richtextbox" => "RichTextBox",
        "picturebox" => "PictureBox", "menustrip" => "MenuStrip",
        "toolstrip" => "ToolStrip", "statusstrip" => "StatusStrip",
        "splitcontainer" => "SplitContainer",
        "flowlayoutpanel" => "FlowLayoutPanel",
        "tablelayoutpanel" => "TableLayoutPanel",
        "linklabel" => "LinkLabel", "maskedtextbox" => "MaskedTextBox",
        "listview" => "ListView", "webbrowser" => "WebBrowser",
        "monthcalendar" => "MonthCalendar",
        "contextmenustrip" => "ContextMenuStrip",
        "timer" => "Timer", "bindingsource" => "BindingSource",
        "tooltip" => "ToolTip", "imagelist" => "ImageList",
        "hscrollbar" => "HScrollBar", "vscrollbar" => "VScrollBar",
        "bindingnavigator" => "BindingNavigator",
        "dataset" => "DataSet", "datatable" => "DataTable",
        "dataadapter" => "DataAdapter",
        "treeview" => "TreeView", "notifyicon" => "NotifyIcon",
        "errorprovider" => "ErrorProvider", "helpprovider" => "HelpProvider",
        "backgroundworker" => "BackgroundWorker",
        "openfiledialog" => "OpenFileDialog",
        "savefiledialog" => "SaveFileDialog",
        "folderbrowserdialog" => "FolderBrowserDialog",
        "colordialog" => "ColorDialog", "fontdialog" => "FontDialog",
        _ => return name.to_string(),
    }.to_string()
}
