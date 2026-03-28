use std::rc::Rc;
use vybe_bytecode::{Value, Op};
use vybe_parser_basic::ast::*;

use crate::compiler::{Compiler, VarResolution, LoopContext};

impl Compiler {
    pub(crate) fn compile_statement(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::Dim(vars) => {
                for var in vars {
                    if let Some(ref bounds) = var.array_bounds {
                        // Dim arr(N) — VB arrays are 0..N inclusive, so size = N+1
                        if let Some(bound_expr) = bounds.first() {
                            self.compile_expression(bound_expr)?;
                            // Add 1 to get size (VB upper bound is inclusive)
                            self.emit_constant(Value::F64(1.0));
                            self.emit(Op::dyn_add);
                            self.emit(Op::array_new_default);
                        } else {
                            // Dim arr() — empty array
                            self.emit_u16(Op::array_new, 0);
                        }
                    } else if let Some(ref init) = var.initializer {
                        self.compile_expression(init)?;
                    } else {
                        self.emit(Op::null);
                    }
                    let name = var.name.as_str().to_lowercase();
                    if var.array_bounds.is_some() {
                        self.known_arrays.insert(name.clone());
                    }
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
                let member_lower = member.as_str().to_lowercase();
                let idx = self.add_string_constant(&member_lower);
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop);
                // Emit controlSetProperty side effect for Me.property or classField.property
                // inside a class. (controlSetProperty is always available as no-op in non-GUI)
                let emit_side_effect = if matches!(*object, Expression::Me) {
                    self.current_scope().resolve_local("me").is_some()
                } else if let Expression::Variable(ref name) = *object {
                    // If the variable is a class field, emit side effect
                    self.class_fields.contains(&name.as_str().to_lowercase())
                        && self.current_scope().resolve_local("me").is_some()
                } else {
                    false
                };
                if emit_side_effect {
                    // Push the object again for the host call
                    self.compile_expression(object)?;
                    let cap = {
                        let mut c = member_lower.chars();
                        match c.next() {
                            None => String::new(),
                            Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
                        }
                    };
                    self.emit_constant(Value::String(Rc::from(cap.as_str())));
                    self.compile_expression(value)?;
                    let set_idx = self.import("vybe:gui", "controlSetProperty");
                    self.emit_host_call(set_idx, 3);
                    self.emit(Op::drop);
                }
            }
            Statement::ArrayAssignment { array, indices, value } => {
                // Stack order for array_set: obj (bottom), key, val (top)
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
                self.compile_expression(value)?;
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
                self.compile_expression(start)?;
                let var_name = variable.as_str().to_lowercase();
                let i_slot = self.define_local(&var_name);
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.compile_expression(end)?;
                self.emit(Op::dyn_le);
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                self.loop_stack.push(LoopContext { start: loop_start, break_jumps: vec![], continue_jumps: vec![] });
                for s in body { self.compile_statement(s)?; }
                let ctx = self.loop_stack.pop().unwrap();
                // Patch continue jumps to step
                let step_offset = self.current_offset();
                for cj in &ctx.continue_jumps { self.patch_jump(*cj); }
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
                for bj in &ctx.break_jumps { self.patch_jump(*bj); }
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
                self.emit(Op::str_length); // works for both arrays and strings
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                let var_name = variable.as_str().to_lowercase();
                let elem_slot = self.define_local(&var_name);
                self.emit_u16(Op::local_set, elem_slot); self.emit(Op::drop);
                self.loop_stack.push(LoopContext { start: loop_start, break_jumps: vec![], continue_jumps: vec![] });
                for s in body { self.compile_statement(s)?; }
                let ctx = self.loop_stack.pop().unwrap();
                for cj in &ctx.continue_jumps { self.patch_jump(*cj); }
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                for bj in &ctx.break_jumps { self.patch_jump(*bj); }
                self.current_scope_mut().end_scope();
            }
            Statement::While { condition, body } => {
                let loop_start = self.current_offset();
                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                self.loop_stack.push(LoopContext { start: loop_start, break_jumps: vec![], continue_jumps: vec![] });
                for s in body { self.compile_statement(s)?; }
                let ctx = self.loop_stack.pop().unwrap();
                for cj in &ctx.continue_jumps { self.patch_jump(*cj); }
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                for bj in &ctx.break_jumps { self.patch_jump(*bj); }
            }
            Statement::DoLoop { pre_condition, body, post_condition } => {
                let loop_start = self.current_offset();
                let mut exit_jump = None;
                if let Some((cond_type, cond_expr)) = pre_condition {
                    self.compile_expression(cond_expr)?;
                    self.emit(Op::dyn_to_bool);
                    exit_jump = Some(match cond_type {
                        LoopConditionType::While => self.emit_jump(Op::br_if_false),
                        LoopConditionType::Until => self.emit_jump(Op::br_if_true),
                    });
                }
                self.loop_stack.push(LoopContext { start: loop_start, break_jumps: vec![], continue_jumps: vec![] });
                for s in body { self.compile_statement(s)?; }
                let ctx = self.loop_stack.pop().unwrap();
                for cj in &ctx.continue_jumps { self.patch_jump(*cj); }
                if let Some((cond_type, cond_expr)) = post_condition {
                    self.compile_expression(cond_expr)?;
                    self.emit(Op::dyn_to_bool);
                    match cond_type {
                        LoopConditionType::While => {
                            let ex = self.emit_jump(Op::br_if_false);
                            self.emit_loop(loop_start);
                            self.patch_jump(ex);
                        }
                        LoopConditionType::Until => {
                            let ex = self.emit_jump(Op::br_if_true);
                            self.emit_loop(loop_start);
                            self.patch_jump(ex);
                        }
                    }
                } else {
                    self.emit_loop(loop_start);
                }
                if let Some(ej) = exit_jump { self.patch_jump(ej); }
                for bj in &ctx.break_jumps { self.patch_jump(*bj); }
            }
            Statement::Call { name, arguments } => {
                let fname = name.as_str().to_lowercase();
                if fname == "console.writeline" || fname == "console.write" || fname == "console" {
                    // Console.WriteLine uses call_import so tests can override wasi:cli log
                    for arg in arguments { self.compile_expression(arg)?; }
                    let idx = self.import("wasi:cli", "log");
                    self.emit_host_call(idx, arguments.len() as u8);
                    self.emit(Op::drop);
                    return Ok(());
                }

                // Me.Show() / Me.Close() inside a class
                if fname.starts_with("me.") {
                    if let Some(me_slot) = self.current_scope().resolve_local("me") {
                        let method = &fname[3..];
                        match method {
                            "show" => {
                                self.emit_u16(Op::local_get, me_slot);
                                let idx = self.import("vybe:gui", "showForm");
                                self.emit_host_call(idx, 1);
                                self.emit(Op::drop);
                                return Ok(());
                            }
                            "close" => {
                                self.emit_u16(Op::local_get, me_slot);
                                let idx = self.import("vybe:gui", "closeForm");
                                self.emit_host_call(idx, 1);
                                self.emit(Op::drop);
                                return Ok(());
                            }
                            _ => {
                                // Me.Method() → resolve as class method
                                if self.class_methods.contains(method) {
                                    self.emit_u16(Op::local_get, me_slot);
                                    let prop_idx = self.add_string_constant(method);
                                    self.emit_u16(Op::struct_get, prop_idx);
                                    self.emit_u16(Op::local_get, me_slot);
                                    for arg in arguments { self.compile_expression(arg)?; }
                                    self.emit_u8(Op::call, (arguments.len() + 1) as u8);
                                    self.emit(Op::drop);
                                    return Ok(());
                                }
                            }
                        }
                    }
                }

                let sig = self.func_signatures.get(&fname).cloned();

                // Inside a class: bare method call → Me.method(Me, args...)
                let is_class_method = self.class_methods.contains(&fname)
                    && self.current_scope().resolve_local("me").is_some();

                if is_class_method {
                    let me_slot = self.current_scope().resolve_local("me").unwrap();
                    // Push the method: Me.methodname
                    self.emit_u16(Op::local_get, me_slot);
                    let prop_idx = self.add_string_constant(&fname);
                    self.emit_u16(Op::struct_get, prop_idx);
                    // Push Me as first arg
                    self.emit_u16(Op::local_get, me_slot);
                    // Push remaining args
                    for arg in arguments { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, (arguments.len() + 1) as u8);
                    self.emit(Op::drop);
                } else {
                    // Push function reference
                    match self.resolve_variable(&fname) {
                        VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                        VarResolution::Global => {
                            let idx = self.add_string_constant(&fname);
                            self.emit_u16(Op::global_get, idx);
                        }
                    }

                    // Compile args — box ByRef params, save box refs for writeback
                    let mut byref_info: Vec<(u16, u16)> = Vec::new(); // (box_local, var_local)
                    for (i, arg) in arguments.iter().enumerate() {
                        let is_byref = sig.as_ref().and_then(|s| s.get(i)).copied().unwrap_or(false);
                        if is_byref {
                            if let Expression::Variable(var) = arg {
                                let var_name = var.as_str().to_lowercase();
                                self.compile_expression(arg)?;
                                self.emit_u16(Op::array_new, 1);
                                let box_local = self.define_local(&format!("__box_{}", i));
                                self.emit(Op::dup);
                                self.emit_u16(Op::local_set, box_local);
                                self.emit(Op::drop);
                                if let VarResolution::Local(var_slot) = self.resolve_variable(&var_name) {
                                    byref_info.push((box_local, var_slot));
                                }
                            } else {
                                self.compile_expression(arg)?;
                                self.emit_u16(Op::array_new, 1);
                            }
                        } else {
                            self.compile_expression(arg)?;
                        }
                    }
                    self.emit_u8(Op::call, arguments.len() as u8);
                    self.emit(Op::drop);

                    // Writeback: read from boxes back into caller's variables
                    for (box_local, var_local) in &byref_info {
                        self.emit_u16(Op::local_get, *box_local);
                        self.emit_constant(Value::F64(0.0));
                        self.emit(Op::array_get);
                        self.emit_u16(Op::local_set, *var_local);
                        self.emit(Op::drop);
                    }
                }
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
            Statement::ExitProperty | Statement::ExitTry => {
                self.emit(Op::null);
                self.emit(Op::r#return);
            }
            Statement::ExitFor | Statement::ExitDo | Statement::ExitWhile | Statement::ExitSelect => {
                let j = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.break_jumps.push(j);
                }
            }
            Statement::Continue(cont_type) => {
                let j = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.continue_jumps.push(j);
                }
            }
            // CompoundAssignment: x += 1, x -= 2, etc.
            Statement::CompoundAssignment { target, operator, value, .. } => {
                let name = target.as_str().to_lowercase();
                match self.resolve_variable(&name) {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(&name);
                        self.emit_u16(Op::global_get, idx);
                    }
                }
                self.compile_expression(value)?;
                match operator {
                    CompoundOp::AddAssign => self.emit(Op::dyn_add),
                    CompoundOp::SubtractAssign => self.emit(Op::f64_sub),
                    CompoundOp::MultiplyAssign => self.emit(Op::f64_mul),
                    CompoundOp::DivideAssign => self.emit(Op::f64_div),
                    CompoundOp::IntDivideAssign => {
                        self.emit(Op::f64_div);
                        let idx = self.import("vybe:math", "floor");
                        self.emit_host_call(idx, 1);
                    }
                    CompoundOp::ConcatAssign => self.emit(Op::str_concat),
                    CompoundOp::ExponentAssign => {
                        let idx = self.import("vybe:math", "pow");
                        self.emit_host_call(idx, 2);
                    }
                    _ => self.emit(Op::dyn_add),
                }
                self.compile_store_ident(target)?;
            }
            // ReDim — resize array
            Statement::ReDim { array, bounds, preserve } => {
                let arr_name = array.as_str().to_lowercase();
                // Get current array
                self.compile_expression(&Expression::Variable(array.clone()))?;
                if let Some(dim) = bounds.first() {
                    self.compile_expression(dim)?;
                } else {
                    self.emit(Op::i32_const_0);
                }
                if *preserve { self.emit(Op::r#true); } else { self.emit(Op::r#false); }
                let idx = self.import("vybe:array", "redim");
                self.emit_host_call(idx, 3);
                // Store result back
                let slot_idx = self.add_string_constant(&arr_name);
                self.emit_u16(Op::global_set, slot_idx);
                self.emit(Op::drop);
            }
            // Using block — resource stored as the named variable
            Statement::Using { variable, resource, body } => {
                self.current_scope_mut().begin_scope();
                self.compile_expression(resource)?;
                let var_name = variable.as_str().to_lowercase();
                let slot = self.define_local(&var_name);
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
                for s in body { self.compile_statement(s)?; }
                self.current_scope_mut().end_scope();
            }
            // SetAssignment: Set x = obj (VB6, same as assignment)
            Statement::SetAssignment { target, value } => {
                self.compile_expression(value)?;
                self.compile_store_ident(target)?;
            }
            // RemoveHandler obj.Event, AddressOf handler
            Statement::RemoveHandler { event_target, handler } => {
                self.emit_constant(Value::String(Rc::from(event_target.as_str())));
                self.emit_constant(Value::String(Rc::from(handler.as_str())));
                let idx = self.import("vybe:gui", "removeHandler");
                self.emit_host_call(idx, 2);
                self.emit(Op::drop);
            }
            // StaticVar — treated as regular local
            Statement::StaticVar { name, initializer, .. } => {
                if let Some(init) = initializer {
                    self.compile_expression(init)?;
                } else {
                    self.emit(Op::null);
                }
                let n = name.as_str().to_lowercase();
                let slot = self.define_local(&n);
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
            }
            // GoTo / Label — compile labels as jump targets, GoTo as jumps
            Statement::Label(_) => {
                // Labels are recorded during a pre-pass; at emit time they're
                // just markers. The label offset is the current position.
                // For now: no-op (labels resolved at statement level within a Sub).
            }
            Statement::GoTo(label) => {
                // GoTo compiles to a forward/backward jump.
                // Without a label pre-pass, we emit a host call that sets a
                // __goto global and returns, allowing the caller to re-dispatch.
                let lbl = self.add_string_constant(&label.as_str().to_lowercase());
                self.emit_u16(Op::r#const, lbl);
                let idx = self.import("vybe:runtime", "goto");
                self.emit_host_call(idx, 1);
                self.emit(Op::drop);
            }
            // On Error GoTo label — wraps remaining code in try/catch
            Statement::OnErrorGoTo(label) => {
                let lbl_name = label.as_str().to_lowercase();
                if lbl_name == "0" {
                    // On Error GoTo 0 — disable error handler (emit try_end)
                    self.emit(Op::try_end);
                } else {
                    // On Error GoTo label — start a try block
                    // The catch will jump to the label (via host "goto" call)
                    let line = self.line;
                    let c = &mut self.chunks[self.current_chunk_idx];
                    c.emit_op(Op::try_start, line);
                    c.emit(0, line); c.emit(0, line); c.emit(0, line); c.emit(0, line);
                }
            }
            // On Error Resume Next — wrap in try/catch that swallows errors
            Statement::OnErrorResumeNext => {
                let line = self.line;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.emit_op(Op::try_start, line);
                c.emit(0, line); c.emit(0, line); c.emit(0, line); c.emit(0, line);
            }
            Statement::Resume(_) => {
                // Resume — VB6 style: continue execution after error
                // In our model this is effectively a no-op after try/catch
            }
            // SyncLock — no-op in single-threaded VM
            Statement::SyncLock { body, .. } => {
                for s in body { self.compile_statement(s)?; }
            }
            // VB6 file I/O — compile as host calls with file number
            Statement::Open { file_path, mode, file_number } => {
                self.compile_expression(file_path)?;
                let mode_str = match mode {
                    FileOpenMode::Input => "Input",
                    FileOpenMode::Output => "Output",
                    FileOpenMode::Append => "Append",
                    FileOpenMode::Binary => "Binary",
                    FileOpenMode::Random => "Random",
                };
                self.emit_constant(Value::String(Rc::from(mode_str)));
                self.compile_expression(file_number)?;
                let idx = self.import("wasi:filesystem", "openFile");
                self.emit_host_call(idx, 3);
                self.emit(Op::drop);
            }
            Statement::CloseFile { file_number } => {
                if let Some(fnum) = file_number {
                    self.compile_expression(fnum)?;
                } else {
                    self.emit_constant(Value::I32(-1)); // close all
                }
                let idx = self.import("wasi:filesystem", "closeFile");
                self.emit_host_call(idx, 1);
                self.emit(Op::drop);
            }
            Statement::PrintFile { file_number, items, newline: _ } => {
                self.compile_expression(file_number)?;
                for v in items { self.compile_expression(v)?; }
                let idx = self.import("wasi:filesystem", "printFile");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::drop);
            }
            Statement::WriteFile { file_number, items } => {
                self.compile_expression(file_number)?;
                for v in items { self.compile_expression(v)?; }
                let idx = self.import("wasi:filesystem", "writeFile");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::drop);
            }
            Statement::InputFile { file_number, variables } => {
                self.compile_expression(file_number)?;
                let idx = self.import("wasi:filesystem", "inputFile");
                self.emit_host_call(idx, 1);
                // Result is an array of values — assign to each variable
                for (i, var) in variables.iter().enumerate() {
                    self.emit(Op::dup);
                    self.emit_constant(Value::F64(i as f64));
                    self.emit(Op::array_get);
                    self.compile_store_ident(var)?;
                }
                self.emit(Op::drop);
            }
            Statement::LineInput { file_number, variable } => {
                self.compile_expression(file_number)?;
                let idx = self.import("wasi:filesystem", "lineInput");
                self.emit_host_call(idx, 1);
                self.compile_store_ident(variable)?;
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
                self.current_scope_mut().begin_scope();
                self.compile_expression(test_expr)?;
                let test_slot = self.define_local("__select_val");
                self.emit_u16(Op::local_set, test_slot);
                self.emit(Op::drop);
                let mut end_jumps = vec![];
                for case in cases {
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
                                    case_true_jump = Some(self.emit_jump(Op::br_if_true));
                                } else {
                                    case_false_jumps.push(self.emit_jump(Op::br_if_false));
                                }
                            }
                            CaseCondition::Range { from, to } => {
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
                    if let Some(tj) = case_true_jump { self.patch_jump(tj); }
                    for s in &case.body { self.compile_statement(s)?; }
                    end_jumps.push(self.emit_jump(Op::br));
                    for fj in case_false_jumps { self.patch_jump(fj); }
                }
                if let Some(els) = else_block {
                    for s in els { self.compile_statement(s)?; }
                }
                for ej in end_jumps { self.patch_jump(ej); }
                self.current_scope_mut().end_scope();
            }
            Statement::With { object, body } => {
                self.current_scope_mut().begin_scope();
                self.compile_expression(object)?;
                let with_slot = self.define_local("__with_obj");
                self.emit_u16(Op::local_set, with_slot);
                self.emit(Op::drop);
                for s in body { self.compile_statement(s)?; }
                self.current_scope_mut().end_scope();
            }
            Statement::AddHandler { event_target, handler } => {
                let parts: Vec<&str> = event_target.splitn(2, '.').collect();
                let control = parts.first().unwrap_or(&"").to_lowercase();
                let event = parts.get(1).unwrap_or(&"Click").to_string();
                match self.resolve_variable(&control) {
                    VarResolution::Local(slot) => {
                        self.emit_u16(Op::local_get, slot);
                        let name_idx = self.add_string_constant("__control_name");
                        self.emit_u16(Op::struct_get, name_idx);
                    }
                    VarResolution::Global => {
                        self.emit_constant(Value::String(Rc::from(control.as_str())));
                    }
                }
                self.emit_constant(Value::String(Rc::from(event.as_str())));
                let handler_lower = handler.to_lowercase();
                let handler_lower = handler_lower.trim_start_matches("me.");
                let idx = self.add_string_constant(handler_lower);
                self.emit_u16(Op::global_get, idx);
                let import_idx = self.import("vybe:gui", "onEvent");
                self.emit_host_call(import_idx, 3);
                self.emit(Op::drop);
            }
            // RaiseEvent EventName(args)
            // → look up __event_EventName on Me, call if not null
            Statement::RaiseEvent { event_name, arguments } => {
                let name_lower = event_name.as_str().to_lowercase();
                let handler_key = format!("__event_{}", name_lower);
                match self.resolve_variable("me") {
                    VarResolution::Local(slot) => {
                        self.emit_u16(Op::local_get, slot);
                        let key_idx = self.add_string_constant(&handler_key);
                        self.emit_u16(Op::struct_get, key_idx);
                        // Check if handler exists (not null)
                        self.emit(Op::dup);
                        self.emit(Op::null);
                        self.emit(Op::dyn_eq);
                        self.emit(Op::dyn_to_bool);
                        let skip = self.emit_jump(Op::br_if_true);
                        // Call handler with args
                        for arg in arguments { self.compile_expression(arg)?; }
                        self.emit_u8(Op::call, arguments.len() as u8);
                        self.emit(Op::drop);
                        let end = self.emit_jump(Op::br);
                        self.patch_jump(skip);
                        self.emit(Op::drop); // drop the null handler
                        self.patch_jump(end);
                    }
                    _ => {}
                }
            }
            _ => {
                // VB statement types not yet compiled
            }
        }
        Ok(())
    }
}
