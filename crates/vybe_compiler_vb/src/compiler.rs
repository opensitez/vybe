use std::rc::Rc;
use std::collections::HashSet;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_basic::ast::*;

use crate::scope::Scope;

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current_chunk_idx: usize,
    line: u32,
    in_method: bool,
    defined_globals: HashSet<String>,
    /// Stack of current function names (lowercase) for VB's "FuncName = value" return convention
    function_name_stack: Vec<String>,
}

impl Compiler {
    pub fn new() -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current_chunk_idx: 0,
            line: 1,
            in_method: false,
            defined_globals: HashSet::new(),
            function_name_stack: Vec::new(),
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        // First pass: register declarations
        for decl in &program.declarations {
            self.compile_declaration(decl)?;
        }
        // Second pass: execute top-level statements
        for stmt in &program.statements {
            self.compile_statement(stmt)?;
        }
        // Auto-call Sub Main() if it was declared (VB entry point convention)
        if self.defined_globals.contains("main") {
            let idx = self.add_string_constant("main");
            self.emit_u16(Op::global_get, idx);
            self.emit_u8(Op::call, 0);
            self.emit(Op::drop);
        }
        self.emit(Op::null);
        self.emit(Op::halt);
        let local_count = self.current_scope().next_slot;
        self.chunks[0].local_count = local_count;
        Ok(self.chunks)
    }

    // -- Import helper --
    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }

    // -- Emit helpers --
    fn emit(&mut self, op: Op) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op(op, line);
    }
    fn emit_u16(&mut self, op: Op, operand: u16) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(op, operand, line);
    }
    fn emit_u8(&mut self, op: Op, operand: u8) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u8(op, operand, line);
    }
    fn emit_constant(&mut self, value: Value) {
        let idx = self.chunks[self.current_chunk_idx].add_constant(value);
        self.emit_u16(Op::r#const, idx);
    }
    fn emit_jump(&mut self, op: Op) -> usize {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_jump(op, line)
    }
    fn patch_jump(&mut self, offset: usize) {
        self.chunks[self.current_chunk_idx].patch_jump(offset);
    }
    fn current_offset(&self) -> usize {
        self.chunks[self.current_chunk_idx].current_offset()
    }
    fn emit_loop(&mut self, target: usize) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_loop(target, line);
    }
    fn add_string_constant(&mut self, s: &str) -> u16 {
        self.chunks[self.current_chunk_idx].add_constant(Value::String(Rc::from(s)))
    }
    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op(Op::call_import, line);
        c.emit((import_idx >> 8) as u8, line);
        c.emit((import_idx & 0xff) as u8, line);
        c.emit(argc, line);
    }
    fn emit_global_set(&mut self, name: &str) {
        let idx = self.add_string_constant(name);
        self.emit_u16(Op::global_set, idx);
        self.defined_globals.insert(name.to_lowercase());
    }

    // -- Scope --
    fn current_scope(&self) -> &Scope { self.scopes.last().unwrap() }
    fn current_scope_mut(&mut self) -> &mut Scope { self.scopes.last_mut().unwrap() }
    fn define_local(&mut self, name: &str) -> u16 { self.current_scope_mut().define_local(name) }

    fn resolve_variable(&self, name: &str) -> VarResolution {
        let lower = name.to_lowercase();
        if let Some(slot) = self.current_scope().resolve_local(&lower) {
            return VarResolution::Local(slot);
        }
        for scope in self.scopes.iter().rev().skip(1) {
            if scope.resolve_local(&lower).is_some() {
                return VarResolution::Local(scope.resolve_local(&lower).unwrap());
            }
        }
        VarResolution::Global
    }

    /// Check if a name refers to a known namespace object (not a user variable).
    fn is_namespace(&self, name: &str) -> bool {
        // Known namespace roots set up by vybe_host::namespaces
        matches!(name,
            "math" | "console" | "convert" | "strings" | "array"
            | "window" | "file" | "io" | "directory"
            | "vybe" | "system" | "application"
        )
    }

    /// Check if an expression is a namespace access chain (e.g. Window.Forms).
    fn is_namespace_expr(&self, expr: &Expression) -> bool {
        match expr {
            Expression::Variable(name) => self.is_namespace(&name.as_str().to_lowercase()),
            Expression::MemberAccess(inner, _) => self.is_namespace_expr(inner),
            _ => false,
        }
    }

    // -- Declarations --
    fn compile_declaration(&mut self, decl: &Declaration) -> Result<(), String> {
        match decl {
            Declaration::Sub(sub) => {
                self.compile_sub(sub)?;
                let name = sub.name.as_str().to_lowercase();
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Function(func) => {
                self.compile_function(func)?;
                let name = func.name.as_str().to_lowercase();
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Class(class) => {
                self.compile_class(class)?;
                let name = class.name.as_str().to_lowercase();
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Variable(vars) => {
                for var in vars {
                    if let Some(ref init) = var.initializer {
                        self.compile_expression(init)?;
                    } else {
                        self.emit(Op::null);
                    }
                    let name = var.name.as_str().to_lowercase();
                    if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                        self.emit_global_set(&name);
                        self.emit(Op::drop);
                    } else {
                        let slot = self.define_local(&name);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                }
            }
            Declaration::Constant(c) => {
                self.compile_expression(&c.value)?;
                let name = c.name.as_str().to_lowercase();
                if self.scopes.len() == 1 {
                    self.emit_global_set(&name);
                    self.emit(Op::drop);
                } else {
                    let slot = self.define_local(&name);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
            }
            Declaration::Imports(_) | Declaration::Namespace(_) |
            Declaration::Enum(_) | Declaration::Interface(_) |
            Declaration::Structure(_) | Declaration::Delegate(_) |
            Declaration::Event(_) => {
                // TODO: implement these
            }
        }
        Ok(())
    }

    // -- Statements --
    fn compile_statement(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::Dim(vars) => {
                for var in vars {
                    if let Some(ref init) = var.initializer {
                        self.compile_expression(init)?;
                    } else {
                        self.emit(Op::null);
                    }
                    let name = var.name.as_str().to_lowercase();
                    let slot = self.define_local(&name);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
            }
            Statement::Assignment { target, value } => {
                self.compile_expression(value)?;
                self.compile_store_ident(target)?;
            }
            Statement::MemberAssignment { object, member, value } => {
                self.compile_expression(object)?;
                self.compile_expression(value)?;
                let idx = self.add_string_constant(&member.as_str().to_lowercase());
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop);
            }
            Statement::ArrayAssignment { array, indices, value } => {
                self.compile_expression(value)?;
                let name = array.as_str().to_lowercase();
                match self.resolve_variable(&name) {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(&name);
                        self.emit_u16(Op::global_get, idx);
                    }
                }
                if let Some(index) = indices.first() {
                    self.compile_expression(index)?;
                }
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            Statement::If { condition, then_branch, elseif_branches, else_branch } => {
                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let else_jump = self.emit_jump(Op::br_if_false);
                for s in then_branch { self.compile_statement(s)?; }
                let mut end_jumps = vec![];
                if !elseif_branches.is_empty() || else_branch.is_some() {
                    end_jumps.push(self.emit_jump(Op::br));
                }
                self.patch_jump(else_jump);
                for (cond, body) in elseif_branches {
                    self.compile_expression(cond)?;
                    self.emit(Op::dyn_to_bool);
                    let next = self.emit_jump(Op::br_if_false);
                    for s in body { self.compile_statement(s)?; }
                    end_jumps.push(self.emit_jump(Op::br));
                    self.patch_jump(next);
                }
                if let Some(els) = else_branch {
                    for s in els { self.compile_statement(s)?; }
                }
                for j in end_jumps { self.patch_jump(j); }
            }
            Statement::For { variable, start, end, step, body } => {
                self.current_scope_mut().begin_scope();
                // i = start
                self.compile_expression(start)?;
                let var_name = variable.as_str().to_lowercase();
                let i_slot = self.define_local(&var_name);
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                // loop
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.compile_expression(end)?;
                self.emit(Op::dyn_le);
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                for s in body { self.compile_statement(s)?; }
                // step
                self.emit_u16(Op::local_get, i_slot);
                if let Some(step_expr) = step {
                    self.compile_expression(step_expr)?;
                } else {
                    self.emit_constant(Value::F64(1.0));
                }
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                self.current_scope_mut().end_scope();
            }
            Statement::ForEach { variable, collection, body } => {
                self.current_scope_mut().begin_scope();
                self.compile_expression(collection)?;
                let arr_slot = self.define_local("__foreach_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__foreach_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                // elem = arr[i]
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                let var_name = variable.as_str().to_lowercase();
                let elem_slot = self.define_local(&var_name);
                self.emit_u16(Op::local_set, elem_slot); self.emit(Op::drop);
                for s in body { self.compile_statement(s)?; }
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                self.current_scope_mut().end_scope();
            }
            Statement::While { condition, body } => {
                let loop_start = self.current_offset();
                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                for s in body { self.compile_statement(s)?; }
                self.emit_loop(loop_start);
                self.patch_jump(exit);
            }
            Statement::DoLoop { pre_condition, body, post_condition } => {
                // Pre-condition: Do While/Until ... Loop
                if let Some((cond_type, cond_expr)) = pre_condition {
                    let loop_start = self.current_offset();
                    self.compile_expression(cond_expr)?;
                    self.emit(Op::dyn_to_bool);
                    let exit = match cond_type {
                        LoopConditionType::While => self.emit_jump(Op::br_if_false),
                        LoopConditionType::Until => self.emit_jump(Op::br_if_true),
                    };
                    for s in body { self.compile_statement(s)?; }
                    self.emit_loop(loop_start);
                    self.patch_jump(exit);
                } else {
                    // Post-condition or infinite: Do ... Loop While/Until
                    let loop_start = self.current_offset();
                    for s in body { self.compile_statement(s)?; }
                    if let Some((cond_type, cond_expr)) = post_condition {
                        self.compile_expression(cond_expr)?;
                        self.emit(Op::dyn_to_bool);
                        match cond_type {
                            LoopConditionType::While => {
                                let exit = self.emit_jump(Op::br_if_false);
                                self.emit_loop(loop_start);
                                self.patch_jump(exit);
                            }
                            LoopConditionType::Until => {
                                let exit = self.emit_jump(Op::br_if_true);
                                self.emit_loop(loop_start);
                                self.patch_jump(exit);
                            }
                        }
                    } else {
                        self.emit_loop(loop_start);
                    }
                }
            }
            Statement::Call { name, arguments } => {
                let fname = name.as_str().to_lowercase();
                // Check for Console.WriteLine
                if fname == "console.writeline" || fname == "console" {
                    for arg in arguments { self.compile_expression(arg)?; }
                    let idx = self.import("wasi:cli", "log");
                    self.emit_host_call(idx, arguments.len() as u8);
                    self.emit(Op::drop);
                    return Ok(());
                }
                // Regular call
                match self.resolve_variable(&fname) {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(&fname);
                        self.emit_u16(Op::global_get, idx);
                    }
                }
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, arguments.len() as u8);
                self.emit(Op::drop);
            }
            Statement::ExpressionStatement(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::drop);
            }
            Statement::Return(Some(expr)) => {
                self.compile_expression(expr)?;
                self.emit(Op::r#return);
            }
            Statement::Return(None) | Statement::ExitSub | Statement::ExitFunction => {
                self.emit(Op::null);
                self.emit(Op::r#return);
            }
            Statement::ExitFor | Statement::ExitDo => {
                // TODO: break from loops
            }
            Statement::Try { body, catches, finally: finally_block } => {
                let try_start_pos = self.current_offset();
                let line = self.line;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.emit_op(Op::try_start, line);
                c.emit(0, line); c.emit(0, line); c.emit(0, line); c.emit(0, line);
                for s in body { self.compile_statement(s)?; }
                self.emit(Op::try_end);
                let skip = self.emit_jump(Op::br);
                // Patch catch offset
                let catch_pos = self.current_offset();
                let ip_after = try_start_pos + 5;
                let catch_offset = catch_pos as i16 - ip_after as i16;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.code[try_start_pos + 1] = (catch_offset >> 8) as u8;
                c.code[try_start_pos + 2] = (catch_offset & 0xff) as u8;
                if let Some(catch) = catches.first() {
                    self.current_scope_mut().begin_scope();
                    if let Some((ref var_name, _)) = catch.variable {
                        let slot = self.define_local(&var_name.as_str().to_lowercase());
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    } else {
                        self.emit(Op::drop);
                    }
                    for s in &catch.body { self.compile_statement(s)?; }
                    self.current_scope_mut().end_scope();
                } else {
                    self.emit(Op::drop);
                }
                self.patch_jump(skip);
                if let Some(fin) = finally_block {
                    for s in fin { self.compile_statement(s)?; }
                }
            }
            Statement::Throw(expr) => {
                if let Some(e) = expr {
                    self.compile_expression(e)?;
                } else {
                    self.emit(Op::null);
                }
                self.emit(Op::throw);
            }
            Statement::Const(c) => {
                self.compile_expression(&c.value)?;
                let name = c.name.as_str().to_lowercase();
                let slot = self.define_local(&name);
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
            }
            Statement::Select { test_expr, cases, else_block } => {
                // Compile the test expression into a temp local
                self.current_scope_mut().begin_scope();
                self.compile_expression(test_expr)?;
                let test_slot = self.define_local("__select_val");
                self.emit_u16(Op::local_set, test_slot);
                self.emit(Op::drop);

                let mut end_jumps = vec![];
                for case in cases {
                    // Each case may have multiple conditions (comma-separated)
                    // Generate: if cond1 OR cond2 OR ... then body
                    let mut case_false_jumps = vec![];
                    let mut case_true_jump = None;

                    for (ci, cond) in case.conditions.iter().enumerate() {
                        match cond {
                            CaseCondition::Value(expr) => {
                                self.emit_u16(Op::local_get, test_slot);
                                self.compile_expression(expr)?;
                                self.emit(Op::dyn_eq);
                                self.emit(Op::dyn_to_bool);
                                if ci < case.conditions.len() - 1 {
                                    // If true, skip to body
                                    case_true_jump = Some(self.emit_jump(Op::br_if_true));
                                } else {
                                    // Last condition: if false, skip body
                                    case_false_jumps.push(self.emit_jump(Op::br_if_false));
                                }
                            }
                            CaseCondition::Range { from, to } => {
                                // test_val >= from AND test_val <= to
                                self.emit_u16(Op::local_get, test_slot);
                                self.compile_expression(from)?;
                                self.emit(Op::dyn_ge);
                                self.emit(Op::dyn_to_bool);
                                let not_in_range = self.emit_jump(Op::br_if_false);
                                self.emit_u16(Op::local_get, test_slot);
                                self.compile_expression(to)?;
                                self.emit(Op::dyn_le);
                                self.emit(Op::dyn_to_bool);
                                if ci < case.conditions.len() - 1 {
                                    case_true_jump = Some(self.emit_jump(Op::br_if_true));
                                    self.patch_jump(not_in_range);
                                } else {
                                    case_false_jumps.push(self.emit_jump(Op::br_if_false));
                                    self.patch_jump(not_in_range);
                                    case_false_jumps.push(self.emit_jump(Op::br));
                                }
                            }
                            CaseCondition::Comparison { op, expr } => {
                                self.emit_u16(Op::local_get, test_slot);
                                self.compile_expression(expr)?;
                                match op {
                                    CompOp::Equal => self.emit(Op::dyn_eq),
                                    CompOp::NotEqual => self.emit(Op::dyn_ne),
                                    CompOp::LessThan => self.emit(Op::dyn_lt),
                                    CompOp::LessThanOrEqual => self.emit(Op::dyn_le),
                                    CompOp::GreaterThan => self.emit(Op::dyn_gt),
                                    CompOp::GreaterThanOrEqual => self.emit(Op::dyn_ge),
                                }
                                self.emit(Op::dyn_to_bool);
                                if ci < case.conditions.len() - 1 {
                                    case_true_jump = Some(self.emit_jump(Op::br_if_true));
                                } else {
                                    case_false_jumps.push(self.emit_jump(Op::br_if_false));
                                }
                            }
                        }
                    }
                    // Patch true jump to here (start of body)
                    if let Some(tj) = case_true_jump {
                        self.patch_jump(tj);
                    }
                    // Body
                    for s in &case.body { self.compile_statement(s)?; }
                    end_jumps.push(self.emit_jump(Op::br));
                    // Patch false jumps to after body
                    for fj in case_false_jumps { self.patch_jump(fj); }
                }
                if let Some(els) = else_block {
                    for s in els { self.compile_statement(s)?; }
                }
                for ej in end_jumps { self.patch_jump(ej); }
                self.current_scope_mut().end_scope();
            }
            Statement::With { object, body } => {
                // Compile `With obj ... End With` by storing obj in a temp local
                self.current_scope_mut().begin_scope();
                self.compile_expression(object)?;
                let with_slot = self.define_local("__with_obj");
                self.emit_u16(Op::local_set, with_slot);
                self.emit(Op::drop);
                for s in body { self.compile_statement(s)?; }
                self.current_scope_mut().end_scope();
            }
            // AddHandler Button1.Click, AddressOf HandleClick
            // → vybe:gui/onEvent("Button1", "Click", handleclick_fn)
            Statement::AddHandler { event_target, handler } => {
                // event_target is "Button1.Click" — split into control + event
                let parts: Vec<&str> = event_target.splitn(2, '.').collect();
                let control = parts.first().unwrap_or(&"").to_lowercase();
                let event = parts.get(1).unwrap_or(&"Click").to_string();

                // Push control name (need to look up __control_name from the variable)
                match self.resolve_variable(&control) {
                    VarResolution::Local(slot) => {
                        // It's a control object — get its __control_name
                        self.emit_u16(Op::local_get, slot);
                        let name_idx = self.add_string_constant("__control_name");
                        self.emit_u16(Op::struct_get, name_idx);
                    }
                    VarResolution::Global => {
                        // Use the name directly as a string
                        self.emit_constant(Value::String(Rc::from(control.as_str())));
                    }
                }
                // Push event name
                self.emit_constant(Value::String(Rc::from(event.as_str())));
                // Push handler function reference
                let handler_lower = handler.to_lowercase();
                let handler_lower = handler_lower.trim_start_matches("me.");
                let idx = self.add_string_constant(handler_lower);
                self.emit_u16(Op::global_get, idx);
                // Call onEvent host function
                let import_idx = self.import("vybe:gui", "onEvent");
                self.emit_host_call(import_idx, 3);
                self.emit(Op::drop);
            }
            _ => {
                // Many VB statement types not yet compiled
            }
        }
        Ok(())
    }

    // -- Expressions --
    fn compile_expression(&mut self, expr: &Expression) -> Result<(), String> {
        match expr {
            Expression::IntegerLiteral(n) => self.emit_constant(Value::F64(*n as f64)),
            Expression::DoubleLiteral(n) => self.emit_constant(Value::F64(*n)),
            Expression::StringLiteral(s) => self.emit_constant(Value::String(Rc::from(s.as_str()))),
            Expression::BooleanLiteral(b) => {
                if *b { self.emit(Op::r#true); } else { self.emit(Op::r#false); }
            }
            Expression::Nothing => self.emit(Op::null),

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

            // Logical
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
                // Check for known VB namespace calls (Console.WriteLine, Math.*, etc.)
                if let Expression::Variable(ref obj_name) = **obj {
                    let obj_lower = obj_name.as_str().to_lowercase();
                    let meth_lower = method.as_str().to_lowercase();
                    let full_name = format!("{}.{}", obj_lower, meth_lower);
                    if let Some(result) = self.try_compile_builtin_method(&full_name, args)? {
                        let _ = result;
                    } else if self.is_namespace(&obj_lower) {
                        // Namespace/static call — no `this`, just get function and call
                        self.compile_expression(obj)?;
                        let prop_idx = self.add_string_constant(&meth_lower);
                        self.emit_u16(Op::struct_get, prop_idx);
                        for arg in args { self.compile_expression(arg)?; }
                        self.emit_u8(Op::call, args.len() as u8);
                    } else {
                        // Instance method call — push this
                        self.compile_expression(obj)?;
                        let prop_idx = self.add_string_constant(&meth_lower);
                        self.emit_u16(Op::struct_get, prop_idx);
                        self.compile_expression(obj)?; // this
                        for arg in args { self.compile_expression(arg)?; }
                        self.emit_u8(Op::call, (args.len() + 1) as u8);
                    }
                } else {
                    // Check for form.Controls.Add(ctrl) pattern
                    if let Expression::MemberAccess(parent, member) = &**obj {
                        let member_lower = member.as_str().to_lowercase();
                        let meth_lower = method.as_str().to_lowercase();
                        if member_lower == "controls" && meth_lower == "add" {
                            // form.Controls.Add(ctrl) → controlsAdd(formName, ctrl)
                            // Get form's __control_name
                            self.compile_expression(parent)?;
                            let name_idx = self.add_string_constant("__control_name");
                            self.emit_u16(Op::struct_get, name_idx);
                            // Push control arg
                            for arg in args { self.compile_expression(arg)?; }
                            let import_idx = self.import("vybe:gui", "controlsAdd");
                            self.emit_host_call(import_idx, (args.len() + 1) as u8);
                            return Ok(());
                        }
                    }

                    // Generic method call — could be namespace chain or instance
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
                        self.compile_expression(obj)?; // this
                        for arg in args { self.compile_expression(arg)?; }
                        self.emit_u8(Op::call, (args.len() + 1) as u8);
                    }
                }
            }
            Expression::New(class_name, args) => {
                let name = class_name.as_str().to_lowercase();
                // Built-in exception types: create an object with message property
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
                    // User-defined class: look up constructor from globals
                    let idx = self.add_string_constant(&name);
                    self.emit_u16(Op::global_get, idx);
                    self.emit_u16(Op::struct_new, 0);
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, (args.len() + 1) as u8);
                }
            }

            // Ternary / If expression
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

            // Cast (simplified — just evaluate the expression)
            Expression::Cast { expr, .. } => {
                self.compile_expression(expr)?;
            }

            // Lambda: Sub() ... End Sub / Function(x) expr
            Expression::Lambda { params, body } => {
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

                match &**body {
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
            }

            // AddressOf — reference to a named function
            Expression::Variable(name) if name.as_str().to_lowercase().starts_with("addressof ") => {
                // Strip "AddressOf " prefix and look up the function
                let func_name = name.as_str()[10..].trim().to_lowercase();
                let idx = self.add_string_constant(&func_name);
                self.emit_u16(Op::global_get, idx);
            }

            _ => {
                // Many VB expression types not yet compiled
                self.emit(Op::null);
            }
        }
        Ok(())
    }

    fn compile_call_expr(&mut self, name: &Identifier, args: &[Expression]) -> Result<(), String> {
        let fname = name.as_str().to_lowercase();

        // VB builtins → host functions
        match fname.as_str() {
            "console.writeline" => {
                for arg in args { self.compile_expression(arg)?; }
                let idx = self.import("wasi:cli", "log");
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            // Type conversion builtins
            "cstr" | "str" | "cint" | "int" | "cdbl" | "cbool" => {
                for arg in args { self.compile_expression(arg)?; }
                let (module, name) = match fname.as_str() {
                    "cstr" | "str" => ("vybe:convert", "toString"),
                    "cint" | "int" => ("vybe:convert", "cint"),
                    "cdbl"         => ("vybe:convert", "cdbl"),
                    "cbool"        => ("vybe:convert", "cbool"),
                    _ => unreachable!(),
                };
                let idx = self.import(module, name);
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            // String builtins
            "len" | "ucase" | "lcase" | "trim" => {
                for arg in args { self.compile_expression(arg)?; }
                let name = match fname.as_str() {
                    "len"   => "length",
                    "ucase" => "ucase",
                    "lcase" => "lcase",
                    "trim"  => "trim",
                    _ => unreachable!(),
                };
                let idx = self.import("vybe:string", name);
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            // Math builtins (function-call form)
            "abs" | "sqr" | "math.floor" | "math.abs" | "math.sqrt" => {
                for arg in args { self.compile_expression(arg)?; }
                let name = match fname.as_str() {
                    "abs" | "math.abs"     => "abs",
                    "sqr" | "math.sqrt"    => "sqrt",
                    "math.floor"           => "floor",
                    _ => unreachable!(),
                };
                let idx = self.import("vybe:math", name);
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            // String functions — all delegate to host
            "left" | "right" | "mid" | "instr" | "replace" | "split" | "join"
            | "ltrim" | "rtrim" | "asc" | "chr" | "space" => {
                for arg in args { self.compile_expression(arg)?; }
                let (module, name) = match fname.as_str() {
                    "left"    => ("vybe:string", "left"),
                    "right"   => ("vybe:string", "right"),
                    "mid"     => ("vybe:string", "mid"),
                    "instr"   => ("vybe:string", "instr"),
                    "replace" => ("vybe:string", "replaceAll"),
                    "split"   => ("vybe:string", "split"),
                    "join"    => ("vybe:array",  "join"),
                    "ltrim"   => ("vybe:string", "ltrim"),
                    "rtrim"   => ("vybe:string", "rtrim"),
                    "asc"     => ("vybe:string", "asc"),
                    "chr"     => ("vybe:string", "chr"),
                    "space"   => ("vybe:string", "space"),
                    _ => unreachable!(),
                };
                let idx = self.import(module, name);
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            // Conversion functions — all delegate to host
            "isnumeric" | "val" | "clng" | "isnothing" => {
                for arg in args { self.compile_expression(arg)?; }
                let (module, name) = match fname.as_str() {
                    "isnumeric" => ("vybe:convert", "isNumeric"),
                    "val"       => ("vybe:convert", "val"),
                    "clng"      => ("vybe:convert", "cint"),
                    "isnothing" => ("vybe:convert", "isNothing"),
                    _ => unreachable!(),
                };
                let idx = self.import(module, name);
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            // Array functions — delegate to host
            "ubound" | "lbound" => {
                for arg in args { self.compile_expression(arg)?; }
                let name = if fname == "ubound" { "ubound" } else { "lbound" };
                let idx = self.import("vybe:array", name);
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            _ => {}
        }

        // Check if name is a local variable — if so, treat as array access
        // (VB uses parens for both calls and array indexing)
        match self.resolve_variable(&fname) {
            VarResolution::Local(slot) => {
                if !self.defined_globals.contains(&fname) {
                    // Local variable — this is array access: arr(index)
                    self.emit_u16(Op::local_get, slot);
                    if let Some(index) = args.first() {
                        self.compile_expression(index)?;
                        self.emit(Op::array_get);
                    }
                    return Ok(());
                }
                self.emit_u16(Op::local_get, slot);
            }
            VarResolution::Global => {
                let idx = self.add_string_constant(&fname);
                self.emit_u16(Op::global_get, idx);
            }
        }
        for arg in args { self.compile_expression(arg)?; }
        self.emit_u8(Op::call, args.len() as u8);
        Ok(())
    }

    /// Try to compile a known VB builtin method call like Console.WriteLine, Math.Floor, etc.
    /// Returns Ok(Some(())) if handled, Ok(None) if not a known builtin.
    fn try_compile_builtin_method(&mut self, full_name: &str, args: &[Expression]) -> Result<Option<()>, String> {
        match full_name {
            // Console
            "console.writeline" => {
                for arg in args { self.compile_expression(arg)?; }
                let idx = self.import("wasi:cli", "log");
                self.emit_host_call(idx, args.len() as u8);
                Ok(Some(()))
            }
            "console.write" => {
                for arg in args { self.compile_expression(arg)?; }
                let idx = self.import("wasi:cli", "log");
                self.emit_host_call(idx, args.len() as u8);
                Ok(Some(()))
            }
            "console.error.writeline" | "console.error" => {
                for arg in args { self.compile_expression(arg)?; }
                let idx = self.import("wasi:cli", "error");
                self.emit_host_call(idx, args.len() as u8);
                Ok(Some(()))
            }
            // Math
            "math.floor" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "floor");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.ceiling" | "math.ceil" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "ceil");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.abs" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "abs");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.sqrt" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "sqrt");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.pow" => {
                self.compile_expression(&args[0])?;
                self.compile_expression(&args[1])?;
                let idx = self.import("vybe:math", "pow");
                self.emit_host_call(idx, 2);
                Ok(Some(()))
            }
            "math.min" => {
                self.compile_expression(&args[0])?;
                self.compile_expression(&args[1])?;
                let idx = self.import("vybe:math", "min");
                self.emit_host_call(idx, 2);
                Ok(Some(()))
            }
            "math.max" => {
                self.compile_expression(&args[0])?;
                self.compile_expression(&args[1])?;
                let idx = self.import("vybe:math", "max");
                self.emit_host_call(idx, 2);
                Ok(Some(()))
            }
            "math.round" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "round");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.sin" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "sin");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.cos" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "cos");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.tan" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "tan");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.log" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "log");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.sign" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "sign");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "math.truncate" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "trunc");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            // String methods
            "string.isnullorempty" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:string", "length");
                self.emit_host_call(idx, 1);
                self.emit_constant(Value::F64(0.0));
                self.emit(Op::dyn_eq);
                Ok(Some(()))
            }
            // Convert
            "convert.toint32" | "convert.toint" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:math", "floor");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "convert.todouble" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:convert", "parseFloat");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            "convert.tostring" => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:convert", "toString");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            // Application.Run
            "application.run" => {
                // Application.Run(form) — pass form's __control_name or string
                if let Some(arg) = args.first() {
                    self.compile_expression(arg)?;
                    // If it's an object, get __control_name; if string, use directly
                } else {
                    self.emit_constant(Value::String(Rc::from("Form1")));
                }
                let idx = self.import("vybe:gui", "runApplication");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            _ => Ok(None),
        }
    }

    fn compile_store_ident(&mut self, target: &Identifier) -> Result<(), String> {
        let name = target.as_str().to_lowercase();
        // VB convention: assigning to the function name sets the return value
        if let Some(func_name) = self.function_name_stack.last() {
            if name == *func_name {
                let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
                self.emit_u16(Op::local_set, rv_slot);
                self.emit(Op::drop);
                return Ok(());
            }
        }
        match self.resolve_variable(&name) {
            VarResolution::Local(slot) => {
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
            }
            VarResolution::Global => {
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
        }
        Ok(())
    }

    fn compile_sub(&mut self, sub: &SubDecl) -> Result<(), String> {
        let name = sub.name.as_str();
        let mut chunk = Chunk::new(name);
        chunk.arity = sub.parameters.len() as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        for param in &sub.parameters { scope.define_local(&param.name.as_str().to_lowercase()); }

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

        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in &upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }
        Ok(())
    }

    fn compile_function(&mut self, func: &FunctionDecl) -> Result<(), String> {
        // Same as Sub but with return value
        self.compile_sub_like(&func.name, &func.parameters, &func.body, Some(&func.name))
    }

    fn compile_sub_like(&mut self, name: &Identifier, params: &[Parameter], body: &[Statement], return_var: Option<&Identifier>) -> Result<(), String> {
        let fname = name.as_str();
        let mut chunk = Chunk::new(fname);
        chunk.arity = params.len() as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        for param in params { scope.define_local(&param.name.as_str().to_lowercase()); }
        // VB functions have an implicit return variable with the function name
        // Use __return_ prefix to avoid shadowing the function's global
        if return_var.is_some() {
            scope.define_local("__return_val");
        }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        // Track function name for VB's "FuncName = value" convention
        if let Some(rv) = return_var {
            self.function_name_stack.push(rv.as_str().to_lowercase());
        }

        // Initialize return variable to null
        if return_var.is_some() {
            self.emit(Op::null);
            let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
            self.emit_u16(Op::local_set, rv_slot);
            self.emit(Op::drop);
        }

        for stmt in body { self.compile_statement(stmt)?; }

        if return_var.is_some() {
            self.function_name_stack.pop();
        }

        // Return the function-name variable (VB convention)
        if return_var.is_some() {
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

        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in &upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }
        Ok(())
    }

    fn compile_class(&mut self, _class: &ClassDecl) -> Result<(), String> {
        // TODO: full class compilation
        self.emit(Op::null);
        Ok(())
    }
}

enum VarResolution { Local(u16), Global }
