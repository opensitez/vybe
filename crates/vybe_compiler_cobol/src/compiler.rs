use std::rc::Rc;
use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_cobol::ast::*;
use vybe_compiler_common as common;

pub struct Compiler {
    chunks: Vec<Chunk>,
    current_chunk_idx: usize,
    line: u32,
    next_local: u16,
    /// Maps data item names (uppercase) to local slots
    data_slots: std::collections::HashMap<String, u16>,
    /// Maps paragraph names to chunk indices
    para_chunks: std::collections::HashMap<String, usize>,
}

impl Compiler {
    pub fn new() -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            current_chunk_idx: 0,
            line: 1,
            next_local: 0,
            data_slots: std::collections::HashMap::new(),
            para_chunks: std::collections::HashMap::new(),
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        // Phase 1: Compile data items (allocate globals with initial values)
        for item in &program.data_items {
            self.compile_data_item(item)?;
        }

        // Phase 2: Compile paragraphs as separate chunks
        for para in &program.paragraphs {
            let ci = self.compile_paragraph(para)?;
            self.para_chunks.insert(para.name.clone(), ci);
        }

        // Phase 2b: Compile classes (OO COBOL 2023)
        for class in &program.classes {
            self.compile_class(class)?;
        }

        // Phase 2c: Compile interfaces
        for iface in &program.interfaces {
            self.compile_interface(iface)?;
        }

        // Phase 3: Compile main body
        for stmt in &program.main_body {
            self.compile_statement(stmt)?;
        }

        self.emit(Op::null);
        self.emit(Op::halt);
        self.chunks[0].local_count = self.next_local;
        common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok(self.chunks)
    }

    // ------------------------------------------------------------------
    // Emit helpers
    // ------------------------------------------------------------------

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

    fn add_string_constant(&mut self, s: &str) -> u16 {
        self.chunks[self.current_chunk_idx].add_constant(Value::String(Rc::from(s)))
    }

    fn current_offset(&self) -> usize {
        self.chunks[self.current_chunk_idx].code.len()
    }

    fn emit_jump(&mut self, op: Op) -> usize {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_jump(op, line)
    }

    fn patch_jump(&mut self, site: usize) {
        self.chunks[self.current_chunk_idx].patch_jump(site);
    }

    fn emit_loop(&mut self, loop_start: usize) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_loop(loop_start, line);
    }

    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }

    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op_u16(Op::call_import, import_idx, line);
        c.emit(argc, line);
    }

    // ------------------------------------------------------------------
    // Data item compilation → globals
    // ------------------------------------------------------------------

    fn compile_data_item(&mut self, item: &DataItem) -> Result<(), String> {
        if item.level == 88 { return Ok(()); } // conditions handled at use site

        let name = item.name.clone();

        if !item.children.is_empty() {
            // Group item → create a struct/dict
            let line = self.line;
            let c = self.current_chunk_idx;
            common::dict::emit_new(&mut self.chunks[c], line);

            // Set fields from children
            for child in &item.children {
                self.emit(Op::dup);
                self.compile_initial_value(&child.pic, &child.value)?;
                let line = self.line;
                common::dict::emit_set_const_key(&mut self.chunks[c], &child.name, line);
                // Recursively handle children of children
                // (for deeply nested groups, we'd need more work)
            }

            let idx = self.add_string_constant(&name);
            self.emit_u16(Op::global_set, idx);
        } else {
            // Elementary item → set global
            self.compile_initial_value(&item.pic, &item.value)?;
            let idx = self.add_string_constant(&name);
            self.emit_u16(Op::global_set, idx);
        }

        // Handle OCCURS → create array
        if let Some(count) = item.occurs {
            let line = self.line;
            let c = self.current_chunk_idx;
            self.emit_u16(Op::array_new, count as u16);
            let idx = self.add_string_constant(&name);
            self.emit_u16(Op::global_set, idx);
        }

        Ok(())
    }

    fn compile_initial_value(&mut self, pic: &Option<String>, value: &Option<Literal>) -> Result<(), String> {
        if let Some(val) = value {
            self.compile_literal(val)?;
        } else if let Some(pic) = pic {
            // Default based on PIC: X → spaces, 9 → 0
            let upper = pic.to_uppercase();
            if upper.starts_with('X') || upper.starts_with('A') {
                self.emit_constant(Value::String(Rc::from("")));
            } else {
                self.emit_constant(Value::F64(0.0));
            }
        } else {
            self.emit(Op::null);
        }
        Ok(())
    }

    fn compile_literal(&mut self, lit: &Literal) -> Result<(), String> {
        match lit {
            Literal::Num(n) => { self.emit_constant(Value::F64(*n)); }
            Literal::Str(s) => { self.emit_constant(Value::String(Rc::from(s.as_str()))); }
            Literal::Spaces => { self.emit_constant(Value::String(Rc::from(" "))); }
            Literal::Zeros => { self.emit_constant(Value::F64(0.0)); }
            Literal::LowValues => { self.emit_constant(Value::String(Rc::from(""))); }
            Literal::HighValues => { self.emit_constant(Value::String(Rc::from("\u{FFFF}"))); }
            Literal::True => { self.emit(Op::r#true); }
            Literal::False => { self.emit(Op::r#false); }
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Paragraph compilation
    // ------------------------------------------------------------------

    fn compile_paragraph(&mut self, para: &Paragraph) -> Result<usize, String> {
        let ci = self.chunks.len();
        let chunk = common::functions::create_function_chunk(&para.name, 0);
        self.chunks.push(chunk);
        let saved = self.current_chunk_idx;
        self.current_chunk_idx = ci;

        for stmt in &para.body {
            self.compile_statement(stmt)?;
        }
        let line = self.line;
        common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
        self.chunks[ci].local_count = 16; // generous local allocation

        self.current_chunk_idx = saved;
        Ok(ci)
    }

    // ------------------------------------------------------------------
    // Statement compilation
    // ------------------------------------------------------------------

    fn compile_statement(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::Display(exprs) => {
                let c = self.current_chunk_idx;
                for expr in exprs {
                    self.compile_expr(expr)?;
                }
                let line = self.line;
                common::io::emit_print(&mut self.chunks[c], exprs.len() as u8, line);
                self.emit(Op::drop);
            }

            Statement::Accept(name) => {
                let i = self.import("wasi:cli", "readLine");
                self.emit_host_call(i, 0);
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_set, idx);
            }

            Statement::Move { src, dsts } => {
                for dst in dsts {
                    self.compile_expr(src)?;
                    let idx = self.add_string_constant(dst);
                    self.emit_u16(Op::global_set, idx);
                }
            }

            Statement::MoveCorresponding { src, dst } => {
                // Copy all properties from src group to dst group
                // Get src dict, get its keys, iterate and copy each to dst
                let src_idx = self.add_string_constant(src);
                let dst_idx = self.add_string_constant(dst);
                self.emit_u16(Op::global_get, src_idx);
                let c = self.current_chunk_idx;
                let line = self.line;
                common::dict::emit_keys(&mut self.chunks[c], line);
                let keys_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, keys_slot);
                // Iterate keys
                self.emit_u16(Op::local_get, keys_slot);
                self.emit(Op::array_length);
                let len_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, len_slot);
                self.emit_constant(Value::I32(0));
                let i_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, i_slot);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit(Op::dyn_lt);
                let exit = self.emit_jump(Op::br_if_false);
                // key = keys[i]
                self.emit_u16(Op::local_get, keys_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                let key_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, key_slot);
                // dst[key] = src[key]
                self.emit_u16(Op::global_get, dst_idx);
                self.emit_u16(Op::global_get, src_idx);
                self.emit_u16(Op::local_get, key_slot);
                self.emit(Op::array_get); // src[key]
                self.emit_u16(Op::local_get, key_slot);
                // struct_set expects [obj, val] with key as constant — use array_set instead
                self.emit(Op::array_set);
                self.emit(Op::drop);
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::I32(1));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
            }

            Statement::Add { srcs, to, giving } => {
                let to_idx = self.add_string_constant(to);
                if let Some(giving_name) = giving {
                    // ADD a b GIVING c → c = a + b
                    let mut first = true;
                    for src in srcs {
                        self.compile_expr(src)?;
                        if !first { self.emit(Op::dyn_add); }
                        first = false;
                    }
                    let idx = self.add_string_constant(giving_name);
                    self.emit_u16(Op::global_set, idx);
                } else {
                    // ADD a TO b → b = b + a
                    self.emit_u16(Op::global_get, to_idx);
                    for src in srcs {
                        self.compile_expr(src)?;
                        self.emit(Op::dyn_add);
                    }
                    self.emit_u16(Op::global_set, to_idx);
                }
            }

            Statement::Subtract { src, from, giving } => {
                let from_idx = self.add_string_constant(from);
                if let Some(giving_name) = giving {
                    self.emit_u16(Op::global_get, from_idx);
                    self.compile_expr(src)?;
                    self.emit(Op::f64_sub);
                    let idx = self.add_string_constant(giving_name);
                    self.emit_u16(Op::global_set, idx);
                } else {
                    self.emit_u16(Op::global_get, from_idx);
                    self.compile_expr(src)?;
                    self.emit(Op::f64_sub);
                    self.emit_u16(Op::global_set, from_idx);
                }
            }

            Statement::Multiply { src, by, giving } => {
                let by_idx = self.add_string_constant(by);
                if let Some(giving_name) = giving {
                    self.compile_expr(src)?;
                    self.emit_u16(Op::global_get, by_idx);
                    self.emit(Op::f64_mul);
                    let idx = self.add_string_constant(giving_name);
                    self.emit_u16(Op::global_set, idx);
                } else {
                    self.emit_u16(Op::global_get, by_idx);
                    self.compile_expr(src)?;
                    self.emit(Op::f64_mul);
                    self.emit_u16(Op::global_set, by_idx);
                }
            }

            Statement::Divide { src, by, giving, remainder } => {
                self.compile_expr(src)?;
                self.compile_expr(by)?;
                self.emit(Op::dup);
                let by_tmp = self.next_local;
                self.next_local += 1;
                self.emit_u16(Op::local_set, by_tmp);
                self.emit(Op::f64_div);
                self.emit(Op::f64_trunc);
                self.emit(Op::dup);
                let gi = self.add_string_constant(giving);
                self.emit_u16(Op::global_set, gi);
                if let Some(rem_name) = remainder {
                    // remainder = src - (quotient * by)
                    self.emit_u16(Op::local_get, by_tmp);
                    self.emit(Op::f64_mul);
                    self.compile_expr(src)?;
                    // swap and subtract
                    let tmp = self.next_local;
                    self.next_local += 1;
                    self.emit_u16(Op::local_set, tmp);
                    self.emit_u16(Op::local_get, tmp);
                    self.emit(Op::f64_sub);
                    // negate (src - q*b, but we have q*b - src currently)
                    self.emit(Op::dyn_neg);
                    let ri = self.add_string_constant(rem_name);
                    self.emit_u16(Op::global_set, ri);
                } else {
                    self.emit(Op::drop);
                }
            }

            Statement::Compute { dst, expr } => {
                self.compile_expr(expr)?;
                let idx = self.add_string_constant(dst);
                self.emit_u16(Op::global_set, idx);
            }

            Statement::If { test, body, else_body } => {
                self.compile_expr(test)?;
                self.emit(Op::dyn_to_bool);
                let mut end_jumps = Vec::new();
                let skip = self.emit_jump(Op::br_if_false);
                for s in body { self.compile_statement(s)?; }
                end_jumps.push(self.emit_jump(Op::br));
                self.patch_jump(skip);
                if let Some(alt) = else_body {
                    for s in alt { self.compile_statement(s)?; }
                }
                for j in end_jumps { self.patch_jump(j); }
            }

            Statement::Evaluate { subject, whens, other } => {
                self.compile_expr(subject)?;
                let disc_slot = self.next_local;
                self.next_local += 1;
                self.emit_u16(Op::local_set, disc_slot);

                let mut end_jumps = Vec::new();

                // For EVALUATE TRUE, subject is Bool(true), comparisons are conditions
                let is_true = matches!(subject, Expr::Bool(true));

                for when in whens {
                    let mut match_jumps = Vec::new();
                    for val in &when.values {
                        if is_true {
                            // EVALUATE TRUE: WHEN condition
                            self.compile_expr(val)?;
                            self.emit(Op::dyn_to_bool);
                        } else {
                            self.emit_u16(Op::local_get, disc_slot);
                            self.compile_expr(val)?;
                            self.emit(Op::dyn_eq);
                        }
                        match_jumps.push(self.emit_jump(Op::br_if_true));
                    }
                    let fail = self.emit_jump(Op::br);
                    for m in &match_jumps { self.patch_jump(*m); }
                    for s in &when.body { self.compile_statement(s)?; }
                    end_jumps.push(self.emit_jump(Op::br));
                    self.patch_jump(fail);
                }
                if let Some(other_body) = other {
                    for s in other_body { self.compile_statement(s)?; }
                }
                for j in end_jumps { self.patch_jump(j); }
            }

            Statement::PerformTimes { count, body } => {
                self.compile_expr(count)?;
                let limit_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, limit_slot);
                self.emit_constant(Value::I32(0));
                let i_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, i_slot);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, limit_slot);
                self.emit(Op::dyn_lt);
                let exit = self.emit_jump(Op::br_if_false);
                for s in body { self.compile_statement(s)?; }
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::I32(1));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
            }

            Statement::PerformUntil { test, body } => {
                let loop_start = self.current_offset();
                self.compile_expr(test)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_true); // exit when condition is true
                for s in body { self.compile_statement(s)?; }
                self.emit_loop(loop_start);
                self.patch_jump(exit);
            }

            Statement::PerformVarying { var, from, by, until, body } => {
                let var_idx = self.add_string_constant(var);
                self.compile_expr(from)?;
                self.emit_u16(Op::global_set, var_idx);
                let loop_start = self.current_offset();
                self.compile_expr(until)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_true);
                for s in body { self.compile_statement(s)?; }
                // Increment: var = var + by
                self.emit_u16(Op::global_get, var_idx);
                self.compile_expr(by)?;
                self.emit(Op::dyn_add);
                self.emit_u16(Op::global_set, var_idx);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
            }

            Statement::PerformParagraph(name) => {
                if let Some(&ci) = self.para_chunks.get(name) {
                    let line = self.line;
                    common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
                    self.emit_u8(Op::call_ref, 0);
                    self.emit(Op::drop);
                } else {
                    // Paragraph not yet compiled — emit as global call
                    let idx = self.add_string_constant(name);
                    self.emit_u16(Op::global_get, idx);
                    self.emit_u8(Op::call_ref, 0);
                    self.emit(Op::drop);
                }
            }

            Statement::StringConcat { sources, into } => {
                let mut first = true;
                for source in sources {
                    self.compile_expr(&source.value)?;
                    if !first { self.emit(Op::str_concat); }
                    first = false;
                }
                if first { self.emit_constant(Value::String(Rc::from(""))); }
                let idx = self.add_string_constant(into);
                self.emit_u16(Op::global_set, idx);
            }

            Statement::Unstring { src, delimiters, into } => {
                let src_idx = self.add_string_constant(src);
                self.emit_u16(Op::global_get, src_idx);
                if let Some(delim) = delimiters.first() {
                    self.emit_constant(Value::String(Rc::from(delim.as_str())));
                } else {
                    self.emit_constant(Value::String(Rc::from(" ")));
                }
                self.emit(Op::str_split);
                // Assign each part to the target variables
                let arr_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, arr_slot);
                for (i, name) in into.iter().enumerate() {
                    self.emit_u16(Op::local_get, arr_slot);
                    self.emit_constant(Value::I32(i as i32));
                    self.emit(Op::array_get);
                    let idx = self.add_string_constant(name);
                    self.emit_u16(Op::global_set, idx);
                }
            }

            Statement::InspectTallying { var, counter, mode: _, target } => {
                // Count occurrences of target in var
                let var_idx = self.add_string_constant(var);
                self.emit_u16(Op::global_get, var_idx);
                self.emit_constant(Value::String(Rc::from(target.as_str())));
                self.emit(Op::str_split);
                self.emit(Op::array_length);
                self.emit_constant(Value::I32(1));
                self.emit(Op::f64_sub);
                let cnt_idx = self.add_string_constant(counter);
                self.emit_u16(Op::global_set, cnt_idx);
            }

            Statement::InspectReplacing { var, mode: _, old, new } => {
                let var_idx = self.add_string_constant(var);
                self.emit_u16(Op::global_get, var_idx);
                self.emit_constant(Value::String(Rc::from(old.as_str())));
                self.emit_constant(Value::String(Rc::from(new.as_str())));
                self.emit(Op::str_replace);
                self.emit_u16(Op::global_set, var_idx);
            }

            Statement::Call { name, args } => {
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_get, idx);
                for arg in args {
                    let ai = self.add_string_constant(arg);
                    self.emit_u16(Op::global_get, ai);
                }
                self.emit_u8(Op::call_ref, args.len() as u8);
                self.emit(Op::drop);
            }

            Statement::Initialize(name) => {
                // Reset to default values (spaces for alpha, zeros for numeric)
                self.emit_constant(Value::String(Rc::from("")));
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_set, idx);
            }

            Statement::Set { target, value } => {
                // SET condition TO TRUE → set parent variable to condition's value
                if *value { self.emit(Op::r#true); } else { self.emit(Op::r#false); }
                let idx = self.add_string_constant(target);
                self.emit_u16(Op::global_set, idx);
            }

            Statement::StopRun => {
                self.emit(Op::null);
                self.emit(Op::halt);
            }

            Statement::Goback => {
                self.emit(Op::null);
                self.emit(Op::r#return);
            }

            Statement::Continue => {
                // No-op
            }

            Statement::GoTo(name) => {
                // Go to paragraph — call it then halt
                if let Some(&ci) = self.para_chunks.get(name) {
                    let line = self.line;
                    common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
                    self.emit_u8(Op::call_ref, 0);
                    self.emit(Op::drop);
                }
            }

            Statement::Raise(expr) => {
                self.compile_expr(expr)?;
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current_chunk_idx], line);
            }

            Statement::JsonGenerate { dst, src } => {
                let src_idx = self.add_string_constant(src);
                self.emit_u16(Op::global_get, src_idx);
                let i = self.import("vybe:json", "stringify");
                self.emit_host_call(i, 1);
                let dst_idx = self.add_string_constant(dst);
                self.emit_u16(Op::global_set, dst_idx);
            }

            Statement::JsonParse { src, dst } => {
                let src_idx = self.add_string_constant(src);
                self.emit_u16(Op::global_get, src_idx);
                let i = self.import("vybe:json", "parse");
                self.emit_host_call(i, 1);
                let dst_idx = self.add_string_constant(dst);
                self.emit_u16(Op::global_set, dst_idx);
            }

            Statement::Open { mode, file } => {
                self.emit_constant(Value::String(Rc::from(file.as_str())));
                let mode_str = match mode {
                    FileMode::Input => "r",
                    FileMode::Output => "w",
                    FileMode::Extend => "a",
                    FileMode::IoMode => "rw",
                };
                self.emit_constant(Value::String(Rc::from(mode_str)));
                let i = self.import("wasi:filesystem", "openFile");
                self.emit_host_call(i, 2);
                let fi = self.add_string_constant(&format!("__file_{}", file));
                self.emit_u16(Op::global_set, fi);
            }

            Statement::Close(file) => {
                let fi = self.add_string_constant(&format!("__file_{}", file));
                self.emit_u16(Op::global_get, fi);
                let i = self.import("wasi:filesystem", "closeFile");
                self.emit_host_call(i, 1);
                self.emit(Op::drop);
            }

            Statement::ReadFile { file, into } => {
                let fi = self.add_string_constant(&format!("__file_{}", file));
                self.emit_u16(Op::global_get, fi);
                let i = self.import("wasi:filesystem", "lineInput");
                self.emit_host_call(i, 1);
                if let Some(var) = into {
                    let idx = self.add_string_constant(var);
                    self.emit_u16(Op::global_set, idx);
                } else {
                    self.emit(Op::drop);
                }
            }

            Statement::WriteFile { record, from } => {
                let fi = self.add_string_constant(&format!("__file_{}", record));
                self.emit_u16(Op::global_get, fi);
                if let Some(var) = from {
                    let vi = self.add_string_constant(var);
                    self.emit_u16(Op::global_get, vi);
                } else {
                    let ri = self.add_string_constant(record);
                    self.emit_u16(Op::global_get, ri);
                }
                let i = self.import("wasi:filesystem", "printFile");
                self.emit_host_call(i, 2);
                self.emit(Op::drop);
            }

            Statement::Sort { file: _, ascending: _, key: _ } => {
                // Sort is complex — simplified: no-op (would need file-level sort)
            }

            Statement::PerformThru { from, thru } => {
                // PERFORM para1 THRU para2 — call each paragraph in sequence
                if let Some(&ci) = self.para_chunks.get(from) {
                    let line = self.line;
                    common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
                    self.emit_u8(Op::call_ref, 0);
                    self.emit(Op::drop);
                }
                if let Some(&ci) = self.para_chunks.get(thru) {
                    let line = self.line;
                    common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
                    self.emit_u8(Op::call_ref, 0);
                    self.emit(Op::drop);
                }
            }

            Statement::SearchTable { table: _, at_end, when_cond, when_body } => {
                // SEARCH — simplified: evaluate condition, if true execute when_body, else at_end
                self.compile_expr(when_cond)?;
                self.emit(Op::dyn_to_bool);
                let skip = self.emit_jump(Op::br_if_false);
                for s in when_body { self.compile_statement(s)?; }
                let end = self.emit_jump(Op::br);
                self.patch_jump(skip);
                for s in at_end { self.compile_statement(s)?; }
                self.patch_jump(end);
            }

            Statement::AcceptFrom { var, source } => {
                match source {
                    AcceptSource::Console => {
                        let i = self.import("wasi:cli", "readLine");
                        self.emit_host_call(i, 0);
                    }
                    AcceptSource::Date => {
                        let i = self.import("wasi:clocks", "toISOString");
                        self.emit_host_call(i, 0);
                    }
                    AcceptSource::Time => {
                        let i = self.import("wasi:clocks", "now");
                        self.emit_host_call(i, 0);
                    }
                    AcceptSource::Day | AcceptSource::DayOfWeek => {
                        let i = self.import("wasi:clocks", "toISOString");
                        self.emit_host_call(i, 0);
                    }
                }
                let idx = self.add_string_constant(var);
                self.emit_u16(Op::global_set, idx);
            }

            Statement::Rewrite { record, from } => {
                // REWRITE record FROM var → write updated record
                let fi = self.add_string_constant(&format!("__file_{}", record));
                self.emit_u16(Op::global_get, fi);
                if let Some(var) = from {
                    let vi = self.add_string_constant(var);
                    self.emit_u16(Op::global_get, vi);
                } else {
                    let ri = self.add_string_constant(record);
                    self.emit_u16(Op::global_get, ri);
                }
                let i = self.import("wasi:filesystem", "printFile");
                self.emit_host_call(i, 2);
                self.emit(Op::drop);
            }

            Statement::DeleteFile(file) => {
                self.emit_constant(Value::String(Rc::from(file.as_str())));
                let i = self.import("wasi:filesystem", "remove");
                self.emit_host_call(i, 1);
                self.emit(Op::drop);
            }

            Statement::StartFile { file, key: _ } => {
                // START positions file pointer — simplified as no-op
                let _ = file;
            }

            Statement::ExitPerform => {
                // EXIT PERFORM → break out of current perform loop
                self.emit(Op::null);
                self.emit(Op::r#return);
            }

            Statement::ExitParagraph => {
                // EXIT PARAGRAPH → return from current paragraph
                self.emit(Op::null);
                self.emit(Op::r#return);
            }

            Statement::Merge { file: _, ascending: _, key: _ } => {
                // MERGE → simplified no-op
            }

            Statement::Copy(_name) => {
                // COPY → preprocessor directive, no-op at compile time
            }

            Statement::InspectConverting { var, from, to } => {
                // INSPECT var CONVERTING from TO to → character-by-character translation
                let var_idx = self.add_string_constant(var);
                self.emit_u16(Op::global_get, var_idx);
                self.emit_constant(Value::String(Rc::from(from.as_str())));
                self.emit_constant(Value::String(Rc::from(to.as_str())));
                self.emit(Op::str_replace);
                self.emit_u16(Op::global_set, var_idx);
            }

            Statement::EvaluateAlso { subjects, whens, other } => {
                // EVALUATE subject1 ALSO subject2 → multi-discriminator switch
                // Simplified: evaluate first subject only
                if let Some(first) = subjects.first() {
                    self.compile_expr(first)?;
                }
                let disc_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, disc_slot);
                let mut end_jumps = Vec::new();
                for when in whens {
                    if let Some(first_vals) = when.values.first() {
                        if let Some(val) = first_vals.first() {
                            self.emit_u16(Op::local_get, disc_slot);
                            self.compile_expr(val)?;
                            self.emit(Op::dyn_eq);
                            let skip = self.emit_jump(Op::br_if_false);
                            for s in &when.body { self.compile_statement(s)?; }
                            end_jumps.push(self.emit_jump(Op::br));
                            self.patch_jump(skip);
                        }
                    }
                }
                if let Some(other_body) = other {
                    for s in other_body { self.compile_statement(s)?; }
                }
                for j in end_jumps { self.patch_jump(j); }
            }

            Statement::Invoke { object, method, args, returning } => {
                // INVOKE object method USING args RETURNING result
                let oi = self.add_string_constant(object);
                self.emit_u16(Op::global_get, oi);
                let mi = self.add_string_constant(method);
                self.emit_u16(Op::struct_get, mi);
                // Push self + args
                self.emit_u16(Op::global_get, oi);
                for arg in args {
                    let ai = self.add_string_constant(arg);
                    self.emit_u16(Op::global_get, ai);
                }
                self.emit_u8(Op::call_ref, (args.len() + 1) as u8);
                if let Some(ret_var) = returning {
                    let ri = self.add_string_constant(ret_var);
                    self.emit_u16(Op::global_set, ri);
                } else {
                    self.emit(Op::drop);
                }
            }

            Statement::TypeDef { name: _, pic: _ } => {
                // TYPEDEF → no runtime effect
            }

            Statement::ValidateStmt(var) => {
                // VALIDATE → check if variable is valid (simplified: no-op)
                let _ = var;
            }

            Statement::FreeStmt(var) => {
                // FREE → set variable to null
                self.emit(Op::null);
                let idx = self.add_string_constant(var);
                self.emit_u16(Op::global_set, idx);
            }

            Statement::AllocateStmt(var) => {
                // ALLOCATE → create empty object
                let line = self.line;
                let c = self.current_chunk_idx;
                common::dict::emit_new(&mut self.chunks[c], line);
                let idx = self.add_string_constant(var);
                self.emit_u16(Op::global_set, idx);
            }

            // ── Async / Threading ──────────────────────────────
            Statement::CallAsync { name, args, handle } => {
                // CALL "program" ASYNC → spawn thread running the program
                // Compile the callee reference
                let ni = self.add_string_constant(name);
                self.emit_u16(Op::global_get, ni);
                // Push args
                for arg in args {
                    let ai = self.add_string_constant(arg);
                    self.emit_u16(Op::global_get, ai);
                }
                // If callee is a function ref, wrap in a closure that calls it with args
                // For now: spawn thread with the function
                let c = self.current_chunk_idx;
                let line = self.line;
                common::threading::emit_thread_spawn(&mut self.chunks[c], line);
                // Store handle
                if let Some(h) = handle {
                    let hi = self.add_string_constant(h);
                    self.emit_u16(Op::global_set, hi);
                } else {
                    self.emit(Op::drop);
                }
            }

            Statement::Wait(handle) => {
                // WAIT FOR handle → join thread, get result
                let hi = self.add_string_constant(handle);
                self.emit_u16(Op::global_get, hi);
                let c = self.current_chunk_idx;
                let line = self.line;
                common::threading::emit_thread_join(&mut self.chunks[c], line);
                self.emit(Op::drop);
            }

            Statement::RunUnit { name, args } => {
                // RUN UNIT → spawn separate thread for program
                let ni = self.add_string_constant(name);
                self.emit_u16(Op::global_get, ni);
                for arg in args {
                    let ai = self.add_string_constant(arg);
                    self.emit_u16(Op::global_get, ai);
                }
                let c = self.current_chunk_idx;
                let line = self.line;
                common::threading::emit_thread_spawn(&mut self.chunks[c], line);
                self.emit(Op::drop);
            }

            Statement::LockMonitor(name) => {
                // LOCK monitor → acquire spinlock
                let ni = self.add_string_constant(name);
                self.emit_u16(Op::global_get, ni);
                let lock_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, lock_slot);
                let c = self.current_chunk_idx;
                let line = self.line;
                common::threading::emit_lock_acquire(&mut self.chunks[c], lock_slot, line);
            }

            Statement::UnlockMonitor(name) => {
                // UNLOCK monitor → release spinlock
                let ni = self.add_string_constant(name);
                self.emit_u16(Op::global_get, ni);
                let lock_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, lock_slot);
                let c = self.current_chunk_idx;
                let line = self.line;
                common::threading::emit_lock_release(&mut self.chunks[c], lock_slot, line);
            }

            Statement::PerformAsync(para_name) => {
                // PERFORM paragraph ASYNC → create fiber (continuation)
                if let Some(&ci) = self.para_chunks.get(para_name) {
                    let line = self.line;
                    common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
                    self.emit(Op::cont_new); // create continuation from function
                    self.emit(Op::null);     // initial value
                    self.emit_u16(Op::resume, 0); // start the fiber
                    self.emit(Op::drop);
                }
            }

            Statement::YieldStmt => {
                // YIELD → suspend current fiber, return control to caller
                self.emit(Op::null);
                self.emit_u16(Op::suspend, 0);
            }

            Statement::SuspendStmt => {
                // SUSPEND → pause execution (same as yield for fibers)
                self.emit(Op::null);
                self.emit_u16(Op::suspend, 0);
            }
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Expression compilation
    // ------------------------------------------------------------------

    fn compile_expr(&mut self, expr: &Expr) -> Result<(), String> {
        match expr {
            Expr::Lit(lit) => self.compile_literal(lit),
            Expr::Bool(b) => {
                if *b { self.emit(Op::r#true); } else { self.emit(Op::r#false); }
                Ok(())
            }
            Expr::Ident(name) => {
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_get, idx);
                Ok(())
            }
            Expr::Subscript(name, index) => {
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_get, idx);
                self.compile_expr(index)?;
                // COBOL is 1-indexed, VM is 0-indexed
                self.emit_constant(Value::I32(1));
                self.emit(Op::f64_sub);
                self.emit(Op::array_get);
                Ok(())
            }
            Expr::Qualified(field, parent) => {
                // X OF Y → Y.X
                let pi = self.add_string_constant(parent);
                self.emit_u16(Op::global_get, pi);
                let fi = self.add_string_constant(field);
                self.emit_u16(Op::struct_get, fi);
                Ok(())
            }
            Expr::BinOp { op, left, right } => {
                self.compile_expr(left)?;
                self.compile_expr(right)?;
                match op {
                    BinOp::Add => { self.emit(Op::dyn_add); }
                    BinOp::Sub => { self.emit(Op::f64_sub); }
                    BinOp::Mul => { self.emit(Op::f64_mul); }
                    BinOp::Div => { self.emit(Op::f64_div); }
                    BinOp::Pow => {
                        let c = self.current_chunk_idx;
                        let line = self.line;
                        common::math::emit_pow(&mut self.chunks[c], line);
                    }
                }
                Ok(())
            }
            Expr::Compare { op, left, right } => {
                self.compile_expr(left)?;
                self.compile_expr(right)?;
                match op {
                    CmpOp::Eq => { self.emit(Op::dyn_eq); }
                    CmpOp::Ne => { self.emit(Op::dyn_ne); }
                    CmpOp::Lt => { self.emit(Op::dyn_lt); }
                    CmpOp::Gt => { self.emit(Op::dyn_gt); }
                    CmpOp::Le => { self.emit(Op::dyn_le); }
                    CmpOp::Ge => { self.emit(Op::dyn_ge); }
                }
                Ok(())
            }
            Expr::Logic { op, left, right } => {
                let c = self.current_chunk_idx;
                let line = self.line;
                match op {
                    LogicOp::And => {
                        self.compile_expr(left)?;
                        let skip = common::expressions::emit_and_start(&mut self.chunks[c], line);
                        self.compile_expr(right)?;
                        common::expressions::emit_short_circuit_end(&mut self.chunks[c], skip);
                    }
                    LogicOp::Or => {
                        self.compile_expr(left)?;
                        let skip = common::expressions::emit_or_start(&mut self.chunks[c], line);
                        self.compile_expr(right)?;
                        common::expressions::emit_short_circuit_end(&mut self.chunks[c], skip);
                    }
                }
                Ok(())
            }
            Expr::Not(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::dyn_not);
                Ok(())
            }
            Expr::FunctionCall { name, args } => {
                self.compile_function(name, args)
            }
            Expr::RefMod { name, start, length } => {
                // Reference modification: name(start:length) → substring
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_get, idx);
                self.compile_expr(start)?;
                // COBOL is 1-indexed, str_substring is 0-indexed
                self.emit_constant(Value::I32(1));
                self.emit(Op::f64_sub);
                if let Some(len) = length {
                    // start + length = end position
                    self.compile_expr(start)?;
                    self.compile_expr(len)?;
                    self.emit(Op::dyn_add);
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::f64_sub);
                } else {
                    self.emit_constant(Value::I32(i32::MAX));
                }
                self.emit(Op::str_substring);
                Ok(())
            }
            Expr::ClassTest { var, class } => {
                self.compile_expr(var)?;
                match class {
                    ClassCondition::Numeric => {
                        let c = self.current_chunk_idx;
                        let line = self.line;
                        common::convert::emit_is_numeric(&mut self.chunks[c], line);
                    }
                    ClassCondition::Alphabetic | ClassCondition::AlphabeticLower | ClassCondition::AlphabeticUpper => {
                        // Simplified: check string length > 0
                        self.emit(Op::str_length);
                        self.emit_constant(Value::I32(0));
                        self.emit(Op::dyn_gt);
                    }
                }
                Ok(())
            }
            Expr::SignTest { var, sign } => {
                self.compile_expr(var)?;
                match sign {
                    SignCondition::Positive => {
                        self.emit_constant(Value::F64(0.0));
                        self.emit(Op::dyn_gt);
                    }
                    SignCondition::Negative => {
                        self.emit_constant(Value::F64(0.0));
                        self.emit(Op::dyn_lt);
                    }
                    SignCondition::Zero => {
                        self.emit_constant(Value::F64(0.0));
                        self.emit(Op::dyn_eq);
                    }
                }
                Ok(())
            }
        }
    }

    // ------------------------------------------------------------------
    // Intrinsic function compilation
    // ------------------------------------------------------------------

    fn compile_function(&mut self, name: &str, args: &[Expr]) -> Result<(), String> {
        let c = self.current_chunk_idx;
        let line = self.line;

        for arg in args { self.compile_expr(arg)?; }

        match name {
            "LENGTH" => { self.emit(Op::str_length); }
            "UPPER-CASE" => { self.emit(Op::str_to_upper); }
            "LOWER-CASE" => { self.emit(Op::str_to_lower); }
            "TRIM" => { self.emit(Op::str_trim); }
            "REVERSE" => { common::collections::emit_reverse(&mut self.chunks[c], line); }
            "CURRENT-DATE" => {
                let i = self.import("wasi:clocks", "toISOString");
                self.emit_host_call(i, 0);
            }
            "MAX" => { common::collections::emit_max(&mut self.chunks[c], args.len() as u8, line); }
            "MIN" => { common::collections::emit_min(&mut self.chunks[c], args.len() as u8, line); }
            "MOD" | "REM" => {
                // MOD(a, b) — args already on stack [a, b]
                self.emit(Op::f64_mod);
            }
            "NUMVAL" | "NUMVAL-C" => {
                common::convert::emit_parse_float(&mut self.chunks[c], line);
            }
            "SUBSTITUTE" => {
                // SUBSTITUTE(str, old, new) → str_replace
                self.emit(Op::str_replace);
            }
            "SQRT" => { common::math::emit_sqrt(&mut self.chunks[c], line); }
            "ABS" => { common::math::emit_abs(&mut self.chunks[c], line); }
            "SUM" => {
                // SUM — args on stack, add them up
                for _ in 1..args.len() { self.emit(Op::dyn_add); }
            }
            "INTEGER" => {
                self.emit(Op::f64_trunc);
            }
            "ORD" => {
                let i = self.import("vybe:string", "charCodeAt");
                self.emit_host_call(i, 1);
            }
            "CHAR" => {
                self.emit(Op::str_from_char_code);
            }
            // Trigonometric
            "SIN" => { common::math::emit_sin(&mut self.chunks[c], line); }
            "COS" => { common::math::emit_cos(&mut self.chunks[c], line); }
            "TAN" => { common::math::emit_tan(&mut self.chunks[c], line); }
            "ASIN" => { let i = self.import("vybe:math", "asin"); self.emit_host_call(i, 1); }
            "ACOS" => { let i = self.import("vybe:math", "acos"); self.emit_host_call(i, 1); }
            "ATAN" => { let i = self.import("vybe:math", "atan"); self.emit_host_call(i, 1); }
            // Logarithmic / exponential
            "LOG" => { common::math::emit_log(&mut self.chunks[c], line); }
            "LOG10" => { let i = self.import("vybe:math", "log10"); self.emit_host_call(i, 1); }
            "EXP" => { common::math::emit_exp(&mut self.chunks[c], line); }
            // Rounding
            "CEILING" => { common::math::emit_ceil(&mut self.chunks[c], line); }
            "FLOOR" => { common::math::emit_floor(&mut self.chunks[c], line); }
            "SIGN" => {
                // Returns -1, 0, or 1
                self.emit(Op::dup);
                self.emit_constant(Value::F64(0.0));
                self.emit(Op::dyn_lt);
                let neg = self.emit_jump(Op::br_if_true);
                self.emit(Op::dup);
                self.emit_constant(Value::F64(0.0));
                self.emit(Op::dyn_gt);
                let pos = self.emit_jump(Op::br_if_true);
                self.emit(Op::drop);
                self.emit_constant(Value::F64(0.0));
                let end1 = self.emit_jump(Op::br);
                self.patch_jump(pos);
                self.emit(Op::drop);
                self.emit_constant(Value::F64(1.0));
                let end2 = self.emit_jump(Op::br);
                self.patch_jump(neg);
                self.emit(Op::drop);
                self.emit_constant(Value::F64(-1.0));
                self.patch_jump(end1);
                self.patch_jump(end2);
            }
            "POWER" => { common::math::emit_pow(&mut self.chunks[c], line); }
            "RANDOM" => { common::math::emit_random(&mut self.chunks[c], line); }
            // Statistical
            "MEAN" => {
                // MEAN(a, b, c) = (a + b + c) / count
                let count = args.len();
                for _ in 1..count { self.emit(Op::dyn_add); }
                self.emit_constant(Value::F64(count as f64));
                self.emit(Op::f64_div);
            }
            "MEDIAN" => {
                // Simplified: return first arg
                for _ in 1..args.len() { self.emit(Op::drop); }
            }
            "VARIANCE" => {
                for _ in 1..args.len() { self.emit(Op::drop); }
                self.emit_constant(Value::F64(0.0));
            }
            // Date functions
            "DATE-OF-INTEGER" | "INTEGER-OF-DATE" => {
                // Simplified pass-through
            }
            "WHEN-COMPILED" => {
                let i = self.import("wasi:clocks", "toISOString");
                self.emit_host_call(i, 0);
            }
            "FORMATTED-DATE" | "FORMATTED-TIME" => {
                let i = self.import("wasi:clocks", "toISOString");
                self.emit_host_call(i, args.len() as u8);
            }
            // String functions
            "CONCATENATE" => {
                for _ in 1..args.len() { self.emit(Op::str_concat); }
            }
            "TEST-NUMVAL" => {
                common::convert::emit_is_numeric(&mut self.chunks[c], line);
            }
            // Financial
            "ANNUITY" | "PRESENT-VALUE" => {
                for _ in 1..args.len() { self.emit(Op::drop); }
            }
            _ => {
                // Unknown function — return null
                for _ in args { self.emit(Op::drop); }
                self.emit(Op::null);
            }
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // OO COBOL — Class compilation using common::classes
    // Same pattern as VB/JS/Ruby/PHP — cross-language compatible
    // ------------------------------------------------------------------

    fn compile_class(&mut self, class: &ClassDef) -> Result<(), String> {
        let class_name = &class.name;
        let parent_name = class.inherits.as_deref().unwrap_or("");

        // Compile data items as class fields
        for item in &class.data_items {
            self.compile_data_item(item)?;
        }

        // Compile instance methods
        let mut method_entries: Vec<(String, usize)> = Vec::new();
        let mut static_method_entries: Vec<(String, usize)> = Vec::new();
        let mut init_chunk: Option<usize> = None;
        let mut init_param_count: u8 = 0;

        for method in &class.instance_methods {
            let ci = self.compile_method(&method)?;
            if method.name.to_uppercase() == "NEW" || method.name.to_uppercase() == "INIT" {
                init_chunk = Some(ci);
                init_param_count = method.params.len() as u8;
            } else {
                method_entries.push((method.name.clone(), ci));
            }
        }

        // Compile factory methods (static)
        for method in &class.factory_methods {
            let ci = self.compile_method(&method)?;
            static_method_entries.push((method.name.clone(), ci));
        }

        // Build constructor chunk (same pattern as Ruby/PHP)
        let ctor_arity = init_param_count;
        let ctor_ci = self.chunks.len();
        let chunk = common::functions::create_function_chunk(
            &format!("{}_ctor", class_name), ctor_arity,
        );
        self.chunks.push(chunk);

        let c = ctor_ci;
        let line = self.line;
        let this_idx = 1u16;

        // Create new typed object
        common::classes::emit_new_typed_object(&mut self.chunks[c], this_idx, class_name, line);
        self.chunks[c].emit_op_u16(Op::local_set, this_idx, line);
        self.chunks[c].local_count = (ctor_arity as u16) + 2;

        // Bind instance methods with cross-language aliases
        for (mname, mci) in &method_entries {
            common::classes::emit_bind_method_with_aliases(
                &mut self.chunks[c], this_idx, mname, *mci, line,
            );
        }

        // Call init if present
        if let Some(init_ci) = init_chunk {
            self.chunks[c].emit_op_u16(Op::local_get, this_idx, line);
            common::functions::emit_ref_func(&mut self.chunks[c], init_ci, 0, line);
            self.chunks[c].emit_op_u16(Op::local_get, this_idx, line);
            for i in 0..ctor_arity {
                self.chunks[c].emit_op_u16(Op::local_get, (i as u16) + 2, line);
            }
            self.chunks[c].emit_op_u8(Op::call_ref, (ctor_arity + 1) as u8, line);
            self.chunks[c].emit_op(Op::drop, line);
        }

        // Return this
        self.chunks[c].emit_op_u16(Op::local_get, this_idx, line);
        self.chunks[c].emit_op(Op::r#return, line);

        // Register type for cross-language compatibility
        let method_names: Vec<String> = method_entries.iter().map(|(n, _)| n.clone()).collect();
        common::classes::register_type(
            &mut self.chunks,
            class_name,
            parent_name,
            method_names,
            method_entries.clone(),
            false,
            class.implements.clone(),
            Some(ctor_ci),
        );

        // Store constructor as global
        let ctor_local = self.next_local; self.next_local += 1;
        let c = self.current_chunk_idx;
        common::classes::emit_store_constructor(
            &mut self.chunks[c], class_name, ctor_ci, ctor_local, line,
        );

        // Bind static methods to constructor
        for (sname, sci) in &static_method_entries {
            let name_idx = self.add_string_constant(class_name);
            self.emit_u16(Op::global_get, name_idx);
            let slot = self.next_local; self.next_local += 1;
            self.emit_u16(Op::local_set, slot);
            common::classes::emit_bind_method_with_aliases(
                &mut self.chunks[self.current_chunk_idx], slot, sname, *sci, line,
            );
        }

        // Inheritance
        if !parent_name.is_empty() {
            let c = self.current_chunk_idx;
            common::classes::emit_inherit_statics(&mut self.chunks[c], parent_name, line);
            let slot = self.next_local; self.next_local += 1;
            common::classes::emit_store_super(&mut self.chunks[c], slot, parent_name, line);
        }

        // Register __new_ClassName function for NEW expression
        let new_name = format!("__new_{}", class_name);
        let ni = self.add_string_constant(&new_name);
        let ci_name = self.add_string_constant(class_name);
        self.emit_u16(Op::global_get, ci_name); // get constructor
        self.emit_u16(Op::global_set, ni);       // store as __new_ClassName

        Ok(())
    }

    fn compile_method(&mut self, method: &MethodDef) -> Result<usize, String> {
        let ci = self.chunks.len();
        let arity = method.params.len() as u8;
        let chunk = common::functions::create_function_chunk(&method.name, arity);
        self.chunks.push(chunk);

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = ci;

        // Compile method body
        for stmt in &method.body {
            self.compile_statement(stmt)?;
        }

        let line = self.line;
        common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
        self.chunks[ci].local_count = (arity as u16) + 16; // generous allocation

        self.current_chunk_idx = saved;
        Ok(ci)
    }

    fn compile_interface(&mut self, iface: &InterfaceDef) -> Result<(), String> {
        // Register interface in type table
        let method_names: Vec<String> = iface.methods.iter().map(|m| m.name.clone()).collect();
        common::classes::register_interface(
            &mut self.chunks,
            &iface.name,
            method_names,
            iface.inherits.clone(),
        );
        Ok(())
    }
}
