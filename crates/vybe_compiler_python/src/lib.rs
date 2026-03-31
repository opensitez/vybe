use std::rc::Rc;
use std::collections::HashMap;
use vybe_parser_python::{Program, Stmt, Expr};
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

/// Minimal line-based Python compiler with small feature set (quick wins).
pub struct Compiler;

impl Compiler {
    pub fn new() -> Self { Compiler }

    pub fn compile(&self, program: &Program) -> Result<Vec<Chunk>, String> {
        // (temporary debug print removed)
    let mut chunk = Chunk::new("<script>");
        // Ensure host logging import is present
        let import_idx = chunk.add_import("wasi:cli", "log");

    // Collect additional chunks (for functions)
    let mut extra_chunks: Vec<Chunk> = Vec::new();

        // Simple locals map: name -> local_index (u16). Local indices start at 1.
        let mut locals: HashMap<String, u16> = HashMap::new();
        let mut max_local: u16 = 0;

        // Helper to allocate a local for a name
        let alloc_local = |name: &str, locals: &mut HashMap<String,u16>, max_local: &mut u16| {
            if let Some(&idx) = locals.get(name) {
                idx
            } else {
                let idx = *max_local + 1; // start at 1
                locals.insert(name.to_string(), idx);
                *max_local = idx;
                idx
            }
        };

        // Helper: compile an expression AST and emit bytecode that pushes its value.
        fn emit_expr(expr: &Expr, chunk: &mut Chunk, locals: &HashMap<String,u16>, pretty_allowed: bool) -> Result<(), String> {
            match expr {
                Expr::Int(n) => { let c = chunk.add_constant(Value::I32(*n)); chunk.emit_op_u16(Op::r#const, c, 0); Ok(()) }
                Expr::Float(f) => { let c = chunk.add_constant(Value::F64(*f)); chunk.emit_op_u16(Op::r#const, c, 0); Ok(()) }
                Expr::Str(s) => { let c = chunk.add_constant(Value::String(Rc::from(s.as_str()))); chunk.emit_op_u16(Op::r#const, c, 0); Ok(()) }
                Expr::Bool(b) => { let c = chunk.add_constant(Value::Bool(*b)); chunk.emit_op_u16(Op::r#const, c, 0); Ok(()) }
                Expr::None => { chunk.emit_op(Op::null, 0); Ok(()) }
                Expr::Ident(name) => {
                    if let Some(&idx) = locals.get(name.as_str()) {
                        chunk.emit_op_u16(Op::local_get, idx, 0);
                        Ok(())
                    } else {
                        Err(format!("Unknown identifier: {}", name))
                    }
                }
                Expr::Unary { op, expr } => {
                    emit_expr(expr, chunk, locals, pretty_allowed)?;
                    match op.as_str() {
                        "not" => { chunk.emit_op(Op::dyn_not, 0); }
                        "-" => { chunk.emit_op(Op::dyn_neg, 0); }
                        _ => {}
                    }
                    Ok(())
                }
                Expr::Binary { op, left, right } => {
                    match op.as_str() {
                        "and" => {
                            // left && right  -> boolean result (short-circuit)
                            emit_expr(left, chunk, locals, false)?;
                            chunk.emit_op(Op::dyn_to_bool, 0);
                            let false_jump = chunk.emit_jump(Op::br_if_false, 0);
                            emit_expr(right, chunk, locals, false)?;
                            chunk.emit_op(Op::dyn_to_bool, 0);
                            let end_jump = chunk.emit_jump(Op::br, 0);
                            chunk.patch_jump(false_jump);
                            chunk.emit_op(Op::r#false, 0);
                            chunk.patch_jump(end_jump);
                        }
                        "or" => {
                            // left || right -> boolean result (short-circuit)
                            emit_expr(left, chunk, locals, false)?;
                            chunk.emit_op(Op::dyn_to_bool, 0);
                            let true_jump = chunk.emit_jump(Op::br_if_true, 0);
                            emit_expr(right, chunk, locals, false)?;
                            chunk.emit_op(Op::dyn_to_bool, 0);
                            let end_jump = chunk.emit_jump(Op::br, 0);
                            chunk.patch_jump(true_jump);
                            chunk.emit_op(Op::r#true, 0);
                            chunk.patch_jump(end_jump);
                        }
                        "==" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::dyn_eq, 0); }
                        "!=" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::dyn_ne, 0); }
                        "<" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::dyn_lt, 0); }
                        ">" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::dyn_gt, 0); }
                        "<=" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::dyn_le, 0); }
                        ">=" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::dyn_ge, 0); }
                        "+" => { emit_expr(left, chunk, locals, pretty_allowed)?; emit_expr(right, chunk, locals, pretty_allowed)?; chunk.emit_op(Op::dyn_add, 0); }
                        "-" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::f64_sub, 0); }
                        "*" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::f64_mul, 0); }
                        "/" => { emit_expr(left, chunk, locals, false)?; emit_expr(right, chunk, locals, false)?; chunk.emit_op(Op::f64_div, 0); }
                        _ => { // fallback: evaluate both
                            emit_expr(left, chunk, locals, false)?;
                            emit_expr(right, chunk, locals, false)?;
                        }
                    }
                    Ok(())
                }
                Expr::List(elems) => {
                    // Try pretty printing if allowed and elements are simple constants
                    let mut can_pretty = pretty_allowed;
                    let mut parts: Vec<String> = Vec::new();
                    if can_pretty {
                        for e in elems.iter() {
                            match e {
                                Expr::Int(n) => parts.push(n.to_string()),
                                Expr::Float(f) => parts.push(f.to_string()),
                                Expr::Str(s) => parts.push(s.clone()),
                                Expr::Bool(b) => parts.push(b.to_string()),
                                Expr::None => parts.push("null".to_string()),
                                _ => { can_pretty = false; break; }
                            }
                        }
                    }
                    if can_pretty && !parts.is_empty() {
                        let s = format!("[{}]", parts.join(", "));
                        let c = chunk.add_constant(Value::String(Rc::from(s)));
                        chunk.emit_op_u16(Op::r#const, c, 0);
                        return Ok(());
                    }
                    for e in elems.iter() { emit_expr(e, chunk, locals, false)?; }
                    chunk.emit_op_u16(Op::array_new, elems.len() as u16, 0);
                    Ok(())
                }
                Expr::Tuple(elems) => {
                    for e in elems.iter() { emit_expr(e, chunk, locals, false)?; }
                    chunk.emit_op_u16(Op::array_new, elems.len() as u16, 0);
                    Ok(())
                }
                Expr::Dict(items) => {
                    // Build a dict via host functions: dictNew() then dictAdd(dict, key, value) for each pair
                    let dict_new_idx = chunk.add_import("vybe:types", "dictNew");
                    let dict_add_idx = chunk.add_import("vybe:types", "dictAdd");
                    // call dictNew() -> pushes a new dict object
                    chunk.emit_op_u16(Op::call_import, dict_new_idx, 0);
                    chunk.emit(0, 0);
                    for (k, v) in items.iter() {
                        // stack: [dict]
                        // duplicate dict so it remains on the stack after the call
                        chunk.emit_op(Op::dup, 0);
                        // stack: [dict, dict]
                        emit_expr(k, chunk, locals, false)?; // push key
                        emit_expr(v, chunk, locals, false)?; // push value
                        // call dictAdd(dict, key, value) -> args count 3
                        chunk.emit_op_u16(Op::call_import, dict_add_idx, 0);
                        chunk.emit(3, 0);
                        // dictAdd returns null — drop the result so the dict remains on stack
                        chunk.emit_op(Op::drop, 0);
                    }
                    Ok(())
                }
                Expr::Index { obj, idx } => {
                    // If the index is a string literal, treat as a dict lookup using host dictItem
                    if let Expr::Str(s) = &**idx {
                        emit_expr(obj, chunk, locals, false)?;
                        let c = chunk.add_constant(Value::String(Rc::from(s.as_str())));
                        chunk.emit_op_u16(Op::r#const, c, 0);
                        let dict_item_idx = chunk.add_import("vybe:types", "dictItem");
                        chunk.emit_op_u16(Op::call_import, dict_item_idx, 0);
                        chunk.emit(2, 0);
                        return Ok(());
                    }
                    emit_expr(obj, chunk, locals, false)?;
                    emit_expr(idx, chunk, locals, false)?;
                    chunk.emit_op(Op::array_get, 0);
                    Ok(())
                }
                Expr::Call { func, args } => {
                    // emit callee then args then call_ref
                    emit_expr(func, chunk, locals, false)?;
                    for a in args.iter() { emit_expr(a, chunk, locals, false)?; }
                    chunk.emit_op_u8(Op::call_ref, args.len() as u8, 0);
                    Ok(())
                }
            }
        }

        // Consume the parsed AST statements
        for stmt in &program.stmts {
            match stmt {
                Stmt::Function { name, args, body } => {
                    // build function chunk
                    let mut fchunk = Chunk::new(name);
                    // map args to locals 1..n in function
                    let mut flocals: HashMap<String,u16> = HashMap::new();
                    let mut fmax: u16 = 0;
                    for a in args.iter().map(|s| s.as_str()).filter(|s| !s.is_empty()) {
                        fmax += 1;
                        flocals.insert(a.to_string(), fmax);
                    }
                    // helper to allocate locals inside function
                    let f_alloc_local = |name: &str, locals: &mut HashMap<String,u16>, max_local: &mut u16| {
                        if let Some(&idx) = locals.get(name) { idx } else { let idx = *max_local + 1; locals.insert(name.to_string(), idx); *max_local = idx; idx }
                    };

                    // compile body statements
                    for s in body.iter() {
                        match s {
                            Stmt::Assign { name, expr } => {
                                let idx = f_alloc_local(name, &mut flocals, &mut fmax);
                                emit_expr(expr, &mut fchunk, &flocals, false)?;
                                fchunk.emit_op_u16(Op::local_set, idx, 0);
                                fchunk.emit_op(Op::drop, 0);
                            }
                            Stmt::If { cond, then_branch, else_branch } => {
                                // emit condition
                                emit_expr(cond, &mut fchunk, &flocals, false)?;
                                // jump to else/end if false
                                let exit_jump = fchunk.emit_jump(Op::br_if_false, 0);
                                // then branch
                                for ts in then_branch.iter() {
                                    match ts {
                                        Stmt::Assign { name, expr } => {
                                            let idx = f_alloc_local(name, &mut flocals, &mut fmax);
                                            emit_expr(expr, &mut fchunk, &flocals, false)?;
                                            fchunk.emit_op_u16(Op::local_set, idx, 0);
                                            fchunk.emit_op(Op::drop, 0);
                                        }
                                        Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut fchunk, &flocals, true)?; } fchunk.emit_op_u16(Op::call_import, import_idx, 0); fchunk.emit(args.len() as u8, 0); }
                                        Stmt::Expr { expr } => {
                                            if let Expr::Ident(s) = expr {
                                                if s == "break" { return Err(format!("Break outside loop in function body")); }
                                                if s == "continue" { return Err(format!("Continue outside loop in function body")); }
                                            }
                                            emit_expr(expr, &mut fchunk, &flocals, false)?;
                                            fchunk.emit_op(Op::drop, 0);
                                        }
                                        Stmt::Return { expr } => { emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op(Op::r#return, 0); }
                                        _ => { return Err(format!("Unsupported stmt in if-then: {:?}", ts)); }
                                    }
                                }
                                // if there's an else, jump past it after then
                                if else_branch.is_some() {
                                    let after_then_jump = fchunk.emit_jump(Op::br, 0);
                                    // patch exit_jump to here (start of else)
                                    fchunk.patch_jump(exit_jump);
                                    if let Some(eb) = else_branch {
                                        for es in eb.iter() {
                                            match es {
                                                Stmt::Assign { name, expr } => { let idx = f_alloc_local(name, &mut flocals, &mut fmax); emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op_u16(Op::local_set, idx, 0); fchunk.emit_op(Op::drop, 0); }
                                                Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut fchunk, &flocals, true)?; } fchunk.emit_op_u16(Op::call_import, import_idx, 0); fchunk.emit(args.len() as u8, 0); }
                                                Stmt::Expr { expr } => { emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op(Op::drop, 0); }
                                                Stmt::Return { expr } => { emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op(Op::r#return, 0); }
                                                Stmt::If { cond, then_branch, else_branch } => {
                                                    // nested if (elif) inside else
                                                    emit_expr(cond, &mut fchunk, &flocals, false)?;
                                                    let exit_jump2 = fchunk.emit_jump(Op::br_if_false, 0);
                                                    for ts2 in then_branch.iter() {
                                                        match ts2 {
                                                            Stmt::Assign { name, expr } => { let idx = f_alloc_local(name, &mut flocals, &mut fmax); emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op_u16(Op::local_set, idx, 0); fchunk.emit_op(Op::drop, 0); }
                                                            Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut fchunk, &flocals, true)?; } fchunk.emit_op_u16(Op::call_import, import_idx, 0); fchunk.emit(args.len() as u8, 0); }
                                                            Stmt::Expr { expr } => {
                                                                if let Expr::Ident(s) = expr {
                                                                    if s == "break" { return Err(format!("Break outside loop in function body")); }
                                                                    if s == "continue" { return Err(format!("Continue outside loop in function body")); }
                                                                }
                                                                emit_expr(expr, &mut fchunk, &flocals, false)?;
                                                                fchunk.emit_op(Op::drop, 0);
                                                            }
                                                            Stmt::Return { expr } => { emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op(Op::r#return, 0); }
                                                            Stmt::Break => { return Err(format!("Break outside loop in function body")); }
                                                            Stmt::Continue => { return Err(format!("Continue outside loop in function body")); }
                                                            _ => { return Err(format!("Unsupported stmt in nested if-then: {:?}", ts2)); }
                                                        }
                                                    }
                                                    if else_branch.is_some() {
                                                        let after_then_jump2 = fchunk.emit_jump(Op::br, 0);
                                                        fchunk.patch_jump(exit_jump2);
                                                        if let Some(eb2) = else_branch {
                                                            for es2 in eb2.iter() {
                                                                match es2 {
                                                                    Stmt::Assign { name, expr } => { let idx = f_alloc_local(name, &mut flocals, &mut fmax); emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op_u16(Op::local_set, idx, 0); fchunk.emit_op(Op::drop, 0); }
                                                                    Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut fchunk, &flocals, true)?; } fchunk.emit_op_u16(Op::call_import, import_idx, 0); fchunk.emit(args.len() as u8, 0); }
                                                                    Stmt::Expr { expr } => { if let Expr::Ident(s) = expr { if s == "break" { return Err(format!("Break outside loop in function body")); } if s == "continue" { return Err(format!("Continue outside loop in function body")); } } emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op(Op::drop, 0); }
                                                                    Stmt::Return { expr } => { emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op(Op::r#return, 0); }
                                                                    Stmt::Break => { return Err(format!("Break outside loop in function body")); }
                                                                    Stmt::Continue => { return Err(format!("Continue outside loop in function body")); }
                                                                    _ => { return Err(format!("Unsupported stmt in nested if-else: {:?}", es2)); }
                                                                }
                                                            }
                                                        }
                                                        fchunk.patch_jump(after_then_jump2);
                                                    } else {
                                                        fchunk.patch_jump(exit_jump2);
                                                    }
                                                }
                                                _ => { return Err(format!("Unsupported stmt in if-else: {:?}", es)); }
                                        }
                                    }
                                    }
                                    fchunk.patch_jump(after_then_jump);
                                } else {
                                    // no else: patch exit_jump to after then
                                    fchunk.patch_jump(exit_jump);
                                }
                            }
                            Stmt::While { cond, body } => {
                                let loop_start = fchunk.current_offset();
                                emit_expr(cond, &mut fchunk, &flocals, false)?;
                                let exit_jump = fchunk.emit_jump(Op::br_if_false, 0);
                                let mut break_jumps: Vec<usize> = Vec::new();
                                for bs in body.iter() {
                                
                                    match bs {
                                        Stmt::Assign { name, expr } => { let idx = f_alloc_local(name, &mut flocals, &mut fmax); emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op_u16(Op::local_set, idx, 0); fchunk.emit_op(Op::drop, 0); }
                                        Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut fchunk, &flocals, true)?; } fchunk.emit_op_u16(Op::call_import, import_idx, 0); fchunk.emit(args.len() as u8, 0); }
                                                                    Stmt::Expr { expr } => { emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op(Op::drop, 0); }
                                        Stmt::Return { expr } => { emit_expr(expr, &mut fchunk, &flocals, false)?; fchunk.emit_op(Op::r#return, 0); }
                                        Stmt::Break => { let j = fchunk.emit_jump(Op::br, 0); break_jumps.push(j); }
                                        Stmt::Continue => { fchunk.emit_loop(loop_start, 0); }
                                        _ => { return Err(format!("Unsupported stmt in function while body: {:?}", bs)); }
                                    }
                                }
                                // patch breaks to after-loop, then jump back to loop start
                                for bj in break_jumps.into_iter() { fchunk.patch_jump(bj); }
                                fchunk.emit_loop(loop_start, 0);
                                fchunk.patch_jump(exit_jump);
                            }
                            Stmt::Print { args } => {
                                for a in args.iter() { emit_expr(a, &mut fchunk, &flocals, true)?; }
                                fchunk.emit_op_u16(Op::call_import, import_idx, 0);
                                fchunk.emit(args.len() as u8, 0);
                            }
                            Stmt::Expr { expr } => {
                                emit_expr(expr, &mut fchunk, &flocals, false)?;
                                fchunk.emit_op(Op::drop, 0);
                            }
                            Stmt::Return { expr } => {
                                emit_expr(expr, &mut fchunk, &flocals, false)?;
                                fchunk.emit_op(Op::r#return, 0);
                            }
                            _ => { return Err(format!("Unsupported stmt in function body: {:?}", s)); }
                        }
                    }

                    // ensure function ends with return (if not already returned)
                    fchunk.emit_op(Op::r#return, 0);

                    // set local count
                    fchunk.local_count = (fmax + 1) as u16;
                    // compute function chunk index: main chunk is 0, extras will follow
                    let func_idx = 1 + extra_chunks.len();
                    // store function chunk for later
                    extra_chunks.push(fchunk);
                    // bind function object to a local in main
                    let idx = alloc_local(name, &mut locals, &mut max_local);
                    chunk.emit_op_u16(Op::ref_func, func_idx as u16, 0);
                    chunk.emit(0, 0);
                    chunk.emit_op_u16(Op::local_set, idx, 0);
                }
                Stmt::Break => { return Err(format!("Break outside loop")); }
                Stmt::Continue => { return Err(format!("Continue outside loop")); }
                Stmt::Assign { name, expr } => {
                    let idx = alloc_local(name, &mut locals, &mut max_local);
                    emit_expr(expr, &mut chunk, &locals, false)?;
                    // store into local (leave value on stack), then drop to not leave value
                    chunk.emit_op_u16(Op::local_set, idx, 0);
                    chunk.emit_op(Op::drop, 0);
                }
                Stmt::If { cond, then_branch, else_branch } => {
                    emit_expr(cond, &mut chunk, &locals, false)?;
                    let exit_jump = chunk.emit_jump(Op::br_if_false, 0);
                    for ts in then_branch.iter() {
                        match ts {
                            Stmt::Assign { name, expr } => { let idx = alloc_local(name, &mut locals, &mut max_local); emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op_u16(Op::local_set, idx, 0); chunk.emit_op(Op::drop, 0); }
                            Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut chunk, &locals, true)?; } chunk.emit_op_u16(Op::call_import, import_idx, 0); chunk.emit(args.len() as u8, 0); }
                            Stmt::Expr { expr } => {
                                if let Expr::Ident(s) = expr {
                                    if s == "break" { return Err(format!("Break outside loop")); }
                                    if s == "continue" { return Err(format!("Continue outside loop")); }
                                }
                                emit_expr(expr, &mut chunk, &locals, false)?;
                                chunk.emit_op(Op::drop, 0);
                            }
                            Stmt::Return { expr } => { return Err(format!("Return outside function: {:?}", expr)); }
                            _ => { return Err(format!("Unsupported stmt in if-then: {:?}", ts)); }
                        }
                    }
                    if else_branch.is_some() {
                        let after_then_jump = chunk.emit_jump(Op::br, 0);
                        chunk.patch_jump(exit_jump);
                            if let Some(eb) = else_branch {
                            for es in eb.iter() {
                                match es {
                                    Stmt::Assign { name, expr } => { let idx = alloc_local(name, &mut locals, &mut max_local); emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op_u16(Op::local_set, idx, 0); chunk.emit_op(Op::drop, 0); }
                                    Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut chunk, &locals, true)?; } chunk.emit_op_u16(Op::call_import, import_idx, 0); chunk.emit(args.len() as u8, 0); }
                                    Stmt::Expr { expr } => { emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op(Op::drop, 0); }
                                    Stmt::Return { expr } => { return Err(format!("Return outside function: {:?}", expr)); }
                                    Stmt::If { cond, then_branch, else_branch } => {
                                        // nested if (elif) inside else
                                        emit_expr(cond, &mut chunk, &locals, false)?;
                                        let exit_jump2 = chunk.emit_jump(Op::br_if_false, 0);
                                        for ts2 in then_branch.iter() {
                                            match ts2 {
                                                Stmt::Assign { name, expr } => { let idx = alloc_local(name, &mut locals, &mut max_local); emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op_u16(Op::local_set, idx, 0); chunk.emit_op(Op::drop, 0); }
                                                Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut chunk, &locals, true)?; } chunk.emit_op_u16(Op::call_import, import_idx, 0); chunk.emit(args.len() as u8, 0); }
                                                Stmt::Expr { expr } => { if let Expr::Ident(s) = expr { if s == "break" { return Err(format!("Break outside loop")); } if s == "continue" { return Err(format!("Continue outside loop")); } } emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op(Op::drop, 0); }
                                                Stmt::Return { expr } => { return Err(format!("Return outside function: {:?}", expr)); }
                                                Stmt::Break => { return Err(format!("Break outside loop")); }
                                                Stmt::Continue => { return Err(format!("Continue outside loop")); }
                                                _ => { return Err(format!("Unsupported stmt in nested if-then: {:?}", ts2)); }
                                            }
                                        }
                                        if else_branch.is_some() {
                                            let after_then_jump2 = chunk.emit_jump(Op::br, 0);
                                            chunk.patch_jump(exit_jump2);
                                            if let Some(eb2) = else_branch {
                                                for es2 in eb2.iter() {
                                                        match es2 {
                                                            Stmt::Assign { name, expr } => { let idx = alloc_local(name, &mut locals, &mut max_local); emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op_u16(Op::local_set, idx, 0); chunk.emit_op(Op::drop, 0); }
                                                            Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut chunk, &locals, true)?; } chunk.emit_op_u16(Op::call_import, import_idx, 0); chunk.emit(args.len() as u8, 0); }
                                                            Stmt::Expr { expr } => { if let Expr::Ident(s) = expr { if s == "break" { return Err(format!("Break outside loop")); } if s == "continue" { return Err(format!("Continue outside loop")); } } emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op(Op::drop, 0); }
                                                            Stmt::Return { expr } => { return Err(format!("Return outside function: {:?}", expr)); }
                                                            Stmt::Break => { return Err(format!("Break outside loop")); }
                                                            Stmt::Continue => { return Err(format!("Continue outside loop")); }
                                                            _ => { return Err(format!("Unsupported stmt in nested if-else: {:?}", es2)); }
                                                        }
                                                }
                                            }
                                            chunk.patch_jump(after_then_jump2);
                                        } else {
                                            chunk.patch_jump(exit_jump2);
                                        }
                                    }
                                    _ => { return Err(format!("Unsupported stmt in if-else: {:?}", es)); }
                                }
                            }
                        }
                        chunk.patch_jump(after_then_jump);
                    } else {
                        chunk.patch_jump(exit_jump);
                    }
                }
                Stmt::While { cond, body } => {
                    let loop_start = chunk.current_offset();
                    emit_expr(cond, &mut chunk, &locals, false)?;
                    let exit_jump = chunk.emit_jump(Op::br_if_false, 0);
                    let mut break_jumps: Vec<usize> = Vec::new();
                    for bs in body.iter() {
                    
                        match bs {
                            Stmt::Assign { name, expr } => { let idx = alloc_local(name, &mut locals, &mut max_local); emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op_u16(Op::local_set, idx, 0); chunk.emit_op(Op::drop, 0); }
                            Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut chunk, &locals, true)?; } chunk.emit_op_u16(Op::call_import, import_idx, 0); chunk.emit(args.len() as u8, 0); }
                            Stmt::Expr { expr } => { emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op(Op::drop, 0); }
                            Stmt::Return { expr } => { return Err(format!("Return outside function: {:?}", expr)); }
                            Stmt::Break => { let j = chunk.emit_jump(Op::br, 0); break_jumps.push(j); }
                            Stmt::Continue => { chunk.emit_loop(loop_start, 0); }
                            Stmt::If { cond, then_branch, else_branch } => {
                                // compile nested if inside while body
                                
                                emit_expr(cond, &mut chunk, &locals, false)?;
                                let exit_jump_if = chunk.emit_jump(Op::br_if_false, 0);
                                for ts in then_branch.iter() {
                                    
                                    match ts {
                                        Stmt::Assign { name, expr } => { let idx = alloc_local(name, &mut locals, &mut max_local); emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op_u16(Op::local_set, idx, 0); chunk.emit_op(Op::drop, 0); }
                                        Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut chunk, &locals, true)?; } chunk.emit_op_u16(Op::call_import, import_idx, 0); chunk.emit(args.len() as u8, 0); }
                                        Stmt::Expr { expr } => {
                                            if let Expr::Ident(s) = expr {
                                                if s == "break" { let j = chunk.emit_jump(Op::br, 0); break_jumps.push(j); }
                                                else if s == "continue" { return Err(format!("Continue outside loop")); }
                                                else { emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op(Op::drop, 0); }
                                            } else { emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op(Op::drop, 0); }
                                        }
                                        Stmt::Return { expr } => { return Err(format!("Return outside function: {:?}", expr)); }
                                        Stmt::Break => { let j = chunk.emit_jump(Op::br, 0); break_jumps.push(j); }
                                        Stmt::Continue => { chunk.emit_loop(loop_start, 0); }
                                        Stmt::If { .. } => { return Err(format!("Nested if inside if in while not supported yet")); }
                                        _ => { return Err(format!("Unsupported stmt in nested if-then: {:?}", ts)); }
                                    }
                                }
                                if else_branch.is_some() {
                                    let after_then_jump = chunk.emit_jump(Op::br, 0);
                                    chunk.patch_jump(exit_jump_if);
                                    if let Some(eb) = else_branch {
                                        for es in eb.iter() {
                                            
                                            match es {
                                                Stmt::Assign { name, expr } => { let idx = alloc_local(name, &mut locals, &mut max_local); emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op_u16(Op::local_set, idx, 0); chunk.emit_op(Op::drop, 0); }
                                                Stmt::Print { args } => { for a in args.iter() { emit_expr(a, &mut chunk, &locals, true)?; } chunk.emit_op_u16(Op::call_import, import_idx, 0); chunk.emit(args.len() as u8, 0); }
                                                Stmt::Expr { expr } => {
                                                    if let Expr::Ident(s) = expr {
                                                        if s == "break" { let j = chunk.emit_jump(Op::br, 0); break_jumps.push(j); }
                                                        else if s == "continue" { return Err(format!("Continue outside loop")); }
                                                        else { emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op(Op::drop, 0); }
                                                    } else { emit_expr(expr, &mut chunk, &locals, false)?; chunk.emit_op(Op::drop, 0); }
                                                }
                                                Stmt::Return { expr } => { return Err(format!("Return outside function: {:?}", expr)); }
                                                Stmt::Break => { let j = chunk.emit_jump(Op::br, 0); break_jumps.push(j); }
                                                Stmt::Continue => { chunk.emit_loop(loop_start, 0); }
                                                _ => { return Err(format!("Unsupported stmt in nested if-else: {:?}", es)); }
                                            }
                                        }
                                    }
                                    chunk.patch_jump(after_then_jump);
                                } else {
                                    chunk.patch_jump(exit_jump_if);
                                }
                            }
                            _ => { return Err(format!("Unsupported stmt in top-level while body: {:?}", bs)); }
                        }
                    }
                    for bj in break_jumps.into_iter() { chunk.patch_jump(bj); }
                    chunk.emit_loop(loop_start, 0);
                    chunk.patch_jump(exit_jump);
                }
                Stmt::Print { args } => {
                    for a in args.iter() {
                        emit_expr(a, &mut chunk, &locals, true)?;
                    }
                    chunk.emit_op_u16(Op::call_import, import_idx, 0);
                    chunk.emit(args.len() as u8, 0);
                }
                Stmt::Return { expr } => {
                    return Err(format!("Return outside function: {:?}", expr));
                }
                Stmt::Expr { expr } => {
                    if let Ok(()) = emit_expr(expr, &mut chunk, &locals, false) {
                        chunk.emit_op(Op::drop, 0);
                    } else {
                        return Err(format!("Unsupported expression stmt: {:?}", expr));
                    }
                }
            }
        }

        // Set local_count to max_local + 1 (0 reserved)
        chunk.local_count = (max_local + 1) as u16;

        // End with halt
        chunk.emit_op(Op::halt, 0);

        // return main chunk plus any extras (functions)
        let mut all = Vec::new();
        all.push(chunk);
        for c in extra_chunks.into_iter() { all.push(c); }
        Ok(all)
    }
}
