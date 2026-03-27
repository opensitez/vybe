use std::rc::Rc;
use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_basic::ast::*;

use crate::compiler::{Compiler, VarResolution};
use crate::scope::Scope;

impl Compiler {
    pub(crate) fn compile_expression(&mut self, expr: &Expression) -> Result<(), String> {
        match expr {
            Expression::IntegerLiteral(n) => self.emit_constant(Value::F64(*n as f64)),
            Expression::DoubleLiteral(n) => self.emit_constant(Value::F64(*n)),
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
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(&name);
                        self.emit_u16(Op::global_get, idx);
                    }
                }
            }
            Expression::MemberAccess(obj, member) => {
                self.compile_expression(obj)?;
                let idx = self.add_string_constant(&member.as_str().to_lowercase());
                self.emit_u16(Op::struct_get, idx);
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
                    self.emit(Op::array_get);
                }
            }
            Expression::ArrayLiteral(elems) => {
                for e in elems { self.compile_expression(e)?; }
                self.emit_u16(Op::array_new, elems.len() as u16);
            }

            // Arithmetic
            Expression::Add(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::dyn_add); }
            Expression::Subtract(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::f64_sub); }
            Expression::Multiply(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::f64_mul); }
            Expression::Divide(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::f64_div); }
            Expression::IntegerDivide(a, b) => {
                self.compile_expression(a)?; self.compile_expression(b)?;
                self.emit(Op::f64_div);
                let idx = self.import("vybe:math", "floor");
                self.emit_host_call(idx, 1);
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
                self.emit(Op::dup); self.emit(Op::dyn_to_bool);
                let end = self.emit_jump(Op::br_if_false);
                self.emit(Op::drop);
                self.compile_expression(b)?;
                self.patch_jump(end);
            }
            Expression::Or(a, b) | Expression::OrElse(a, b) => {
                self.compile_expression(a)?;
                self.emit(Op::dup); self.emit(Op::dyn_to_bool);
                let end = self.emit_jump(Op::br_if_true);
                self.emit(Op::drop);
                self.compile_expression(b)?;
                self.patch_jump(end);
            }

            // Bitwise
            Expression::BitShiftLeft(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::i32_shl); }
            Expression::BitShiftRight(a, b) => { self.compile_expression(a)?; self.compile_expression(b)?; self.emit(Op::i32_shr_s); }

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
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);
                self.compile_expression(then_val)?;
                let end_j = self.emit_jump(Op::br);
                self.patch_jump(else_j);
                if let Some(ev) = else_val {
                    self.compile_expression(ev)?;
                } else {
                    self.emit(Op::null);
                }
                self.patch_jump(end_j);
            }

            // Cast
            Expression::Cast { expr, .. } => {
                self.compile_expression(expr)?;
            }

            // Lambda
            Expression::Lambda { params, body } => {
                self.compile_lambda(params, body)?;
            }

            // AddressOf (parsed as Variable with "AddressOf " prefix)
            Expression::Variable(name) if name.as_str().to_lowercase().starts_with("addressof ") => {
                let func_name = name.as_str()[10..].trim().to_lowercase();
                let idx = self.add_string_constant(&func_name);
                self.emit_u16(Op::global_get, idx);
            }

            _ => {
                self.emit(Op::null);
            }
        }
        Ok(())
    }

    fn compile_method_call(&mut self, obj: &Expression, method: &Identifier, args: &[Expression]) -> Result<(), String> {
        // MyBase.New(args) — call parent constructor with Me
        if matches!(obj, Expression::MyBase) && method.as_str().eq_ignore_ascii_case("New") {
            // The parent constructor is stored as __super on Me (set by Inherits compilation)
            // Or we look it up from the class's inherits info
            // For now: MyBase.New(args) → get parent from __super, call with Me + args
            match self.resolve_variable("me") {
                VarResolution::Local(slot) => {
                    self.emit_u16(Op::local_get, slot);
                    let super_idx = self.add_string_constant("__super");
                    self.emit_u16(Op::struct_get, super_idx);
                    // Push Me as first arg
                    self.emit_u16(Op::local_get, slot);
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, (args.len() + 1) as u8);
                    self.emit(Op::drop);
                }
                _ => {}
            }
            return Ok(());
        }

        // MyBase.Method(args) — call parent's method with Me
        if matches!(obj, Expression::MyBase) {
            let meth_lower = method.as_str().to_lowercase();
            match self.resolve_variable("me") {
                VarResolution::Local(slot) => {
                    // Get method from Me (parent attached it)
                    self.emit_u16(Op::local_get, slot);
                    let prop_idx = self.add_string_constant(&meth_lower);
                    self.emit_u16(Op::struct_get, prop_idx);
                    self.emit_u16(Op::local_get, slot); // Me as this
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, (args.len() + 1) as u8);
                }
                _ => { self.emit(Op::null); }
            }
            return Ok(());
        }

        if let Expression::Variable(ref obj_name) = *obj {
            let obj_lower = obj_name.as_str().to_lowercase();
            let meth_lower = method.as_str().to_lowercase();
            let full_name = format!("{}.{}", obj_lower, meth_lower);
            if let Some(result) = self.try_compile_builtin_method(&full_name, args)? {
                let _ = result;
            } else if self.is_namespace(&obj_lower) || self.defined_classes.contains(&obj_lower) {
                // Namespace or class static call — no `this`
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
                    self.compile_expression(parent)?;
                    let name_idx = self.add_string_constant("__control_name");
                    self.emit_u16(Op::struct_get, name_idx);
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
        if name == "exception" || name == "argumentexception" || name == "invalidoperationexception"
            || name == "notimplementedexception" || name == "notsupportedexception" {
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
        } else {
            let idx = self.add_string_constant(&name);
            self.emit_u16(Op::global_get, idx);
            self.emit_u16(Op::struct_new, 0);
            for arg in args { self.compile_expression(arg)?; }
            self.emit_u8(Op::call, (args.len() + 1) as u8);
        }
        Ok(())
    }

    fn compile_lambda(&mut self, params: &[Parameter], body: &LambdaBody) -> Result<(), String> {
        let mut chunk = Chunk::new("<lambda>");
        chunk.arity = params.len() as u8;
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
