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
                // Check if this is a namespace.property that should be auto-called
                // e.g. DateTime.Now, Environment.NewLine, Guid.NewGuid
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

            // Bitwise operators (Xor is both logical and bitwise in VB)
            Expression::Xor(a, b) => {
                self.compile_expression(a)?; self.compile_expression(b)?;
                self.emit(Op::dyn_ne); // XOR: true if different
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

            // TypeOf expr Is Type
            Expression::TypeOf { expr, type_name } => {
                self.compile_expression(expr)?;
                let idx = self.import("vybe:convert", "typeName");
                self.emit_host_call(idx, 1);
                self.emit_constant(Value::String(Rc::from(type_name.as_str())));
                self.emit(Op::dyn_eq);
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

            // Await — suspend fiber until promise resolves (same as JS await)
            Expression::Await(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::r#await);
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
                    self.emit_u16(Op::array_new, items.len() as u16);
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

        // User-defined class takes priority over built-in types
        if self.defined_classes.contains(&name) {
            let idx = self.add_string_constant(&name);
            self.emit_u16(Op::global_get, idx);
            self.emit_u16(Op::struct_new, 0);
            for arg in args { self.compile_expression(arg)?; }
            self.emit_u8(Op::call, (args.len() + 1) as u8);
            return Ok(());
        }

        // Strip generic type params: "dictionary(of string, integer)" → "dictionary"
        let name = if let Some(paren) = name.find("(of ") {
            name[..paren].to_string()
        } else {
            name
        };

        // Built-in types → host constructor (no struct_new needed, host creates the object)
        let host_ctor: Option<(&str, &str)> = match name.as_str() {
            "datetime"      => Some(("vybe:types", "dateTimeNew")),
            "stringbuilder" | "system.text.stringbuilder" => Some(("vybe:types", "stringBuilderNew")),
            "list" | "list(of string)" | "list(of integer)" | "list(of object)" | "list(of double)"
                            => Some(("vybe:types", "listNew")),
            "dictionary" | "dictionary(of string, string)" | "dictionary(of string, object)"
                            => Some(("vybe:types", "dictNew")),
            "queue" | "queue(of string)" | "queue(of integer)" | "queue(of object)"
                            => Some(("vybe:types", "queueNew")),
            "stack" | "stack(of string)" | "stack(of integer)" | "stack(of object)"
                            => Some(("vybe:types", "stackNew")),
            "hashset" | "hashset(of string)" | "hashset(of integer)" | "hashset(of object)"
                            => Some(("vybe:types", "hashSetNew")),
            "datatable"     => Some(("vybe:data", "dataTableNew")),
            "dataset"       => Some(("vybe:data", "dataSetNew")),
            "point"         => Some(("vybe:drawing", "pointNew")),
            "size"          => Some(("vybe:drawing", "sizeNew")),
            "font"          => Some(("vybe:drawing", "fontNew")),
            "random"        => Some(("vybe:threading", "randomNew")),
            "stopwatch"     => Some(("vybe:threading", "stopwatchNew")),
            "arraylist" | "system.collections.arraylist" => Some(("vybe:types", "listNew")),
            "hashtable" | "system.collections.hashtable" => Some(("vybe:types", "dictNew")),
            "collection" | "system.collections.collection" => Some(("vybe:types", "listNew")),
            "sortedlist"    => Some(("vybe:types", "dictNew")),
            "streamreader" | "system.io.streamreader" => Some(("vybe:net", "streamReaderNew")),
            "streamwriter" | "system.io.streamwriter" => Some(("vybe:net", "streamWriterNew")),
            // Database
            "sqlconnection" | "system.data.sqlclient.sqlconnection" => Some(("vybe:database", "connect")),
            "sqlcommand" | "system.data.sqlclient.sqlcommand" => Some(("vybe:data", "dataTableNew")),  // simplified
            "sqldataadapter" | "system.data.sqlclient.sqldataadapter" => Some(("vybe:data", "dataTableNew")),
            "oledbconnection" | "system.data.oledb.oledbconnection" => Some(("vybe:database", "connect")),
            "oledbcommand" | "system.data.oledb.oledbcommand" => Some(("vybe:data", "dataTableNew")),
            "adodb.connection" => Some(("vybe:database", "connect")),
            "adodb.recordset" | "adodb.command" => Some(("vybe:data", "dataTableNew")),
            // Network
            "tcpclient" | "system.net.sockets.tcpclient" => Some(("vybe:net", "tcpConnect")),
            "tcplistener" | "system.net.sockets.tcplistener" => Some(("vybe:net", "tcpListenerNew")),
            "udpclient" | "system.net.sockets.udpclient" => Some(("vybe:net", "udpNew")),
            // WinForms controls — use Window.Forms namespace constructors
            "form" => Some(("vybe:gui", "newForm")),
            "toolstripstatuslabel" | "toolstripmenuitem" | "toolstriptextbox"
            | "toolstripcombobox" | "toolstripseparator" => Some(("vybe:types", "listNew")), // simplified
            _ => None,
        };

        if let Some((module, fn_name)) = host_ctor {
            self.emit(Op::null);
            for arg in args { self.compile_expression(arg)?; }
            let idx = self.import(module, fn_name);
            self.emit_host_call(idx, (args.len() + 1) as u8);
            return Ok(());
        }

        // WinForms control constructors: New Button(), New TextBox(), etc.
        // These map to vybe:gui/new_{CapitalizedName}
        let control_types = [
            "button", "label", "textbox", "checkbox", "radiobutton",
            "combobox", "listbox", "panel", "groupbox", "tabcontrol",
            "tabpage", "datagridview", "progressbar", "trackbar",
            "numericupdown", "datetimepicker", "richtextbox", "picturebox",
            "menustrip", "toolstrip", "statusstrip", "splitcontainer",
            "flowlayoutpanel", "tablelayoutpanel", "linklabel", "maskedtextbox",
            "listview", "webbrowser", "monthcalendar", "contextmenustrip",
            "timer", "bindingsource", "tooltip", "imagelist",
            "openfiledialog", "savefiledialog", "folderbrowserdialog",
            "colordialog", "fontdialog",
        ];
        // Also strip System.Windows.Forms. prefix
        let bare_name = name.strip_prefix("system.windows.forms.").unwrap_or(&name);
        if control_types.contains(&bare_name.as_ref()) {
            // Capitalize: "textbox" → "TextBox" — use the registered host fn name
            // The host fn is registered as "new_TextBox", "new_Button", etc.
            // We look it up by the namespace: global "window" → "forms" → control_name
            self.emit(Op::null);
            for arg in args { self.compile_expression(arg)?; }
            // Find the matching new_X host function
            let capitalized = capitalize_control_name(bare_name);
            let hn = format!("new_{}", capitalized);
            let idx = self.import("vybe:gui", &hn);
            self.emit_host_call(idx, (args.len() + 1) as u8);
            return Ok(());
        }

        // User-defined class: look up constructor from globals
        let idx = self.add_string_constant(&name);
        self.emit_u16(Op::global_get, idx);
        self.emit_u16(Op::struct_new, 0);
        for arg in args { self.compile_expression(arg)?; }
        self.emit_u8(Op::call, (args.len() + 1) as u8);
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

/// Capitalize control type name: "textbox" → "TextBox", "datagridview" → "DataGridView"
fn capitalize_control_name(name: &str) -> String {
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
        "openfiledialog" => "OpenFileDialog",
        "savefiledialog" => "SaveFileDialog",
        "folderbrowserdialog" => "FolderBrowserDialog",
        "colordialog" => "ColorDialog", "fontdialog" => "FontDialog",
        _ => return name.to_string(),
    }.to_string()
}
