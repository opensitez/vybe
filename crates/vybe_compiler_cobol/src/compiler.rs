use std::rc::Rc;
use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_cobol::ast::*;
use vybe_compiler_common as common;
use vybe_compiler_common::collections as common_collections;
use vybe_compiler_common::strings as common_strings;

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

    /// CLS-compliant name normalization for cross-language access.
    /// Converts COBOL names (uppercase, hyphens) to a form other languages can use.
    /// "WS-CUSTOMER-NAME" → "ws_customer_name"
    fn cls_normalize(name: &str) -> String {
        name.to_lowercase().replace('-', "_")
    }

    /// Emit global_set with CLS alias.
    /// Stores the value under both the original COBOL name AND the normalized CLS name.
    /// COBOL code uses: global_get "WS-CUSTOMER-NAME" (original)
    /// VB/C#/JS use:    global_get "ws_customer_name" (CLS alias)
    fn emit_global_set_with_cls(&mut self, name: &str) {
        // Store under original name (for COBOL internal use)
        let idx = self.add_string_constant(name);
        self.emit_u16(Op::global_set, idx);

        // Also store CLS alias if different (for cross-language access)
        // global_set peeks (doesn't pop), so value is still on stack
        let cls_name = Self::cls_normalize(name);
        if cls_name != name {
            let cls_idx = self.add_string_constant(&cls_name);
            self.emit_u16(Op::global_set, cls_idx);
        }
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

            self.emit_global_set_with_cls(&name);
        } else {
            // Elementary item → set global
            self.compile_initial_value(&item.pic, &item.value)?;
            self.emit_global_set_with_cls(&name);

            // Store PIC metadata for formatting (internal)
            if let Some(pic) = &item.pic {
                self.emit_constant(Value::String(Rc::from(pic.as_str())));
                let pk = self.add_string_constant(&format!("__PIC_{}", name));
                self.emit_u16(Op::global_set, pk);
            }
        }

        // Handle OCCURS → create array
        if let Some(count) = item.occurs {
            common_collections::emit_array_new(&mut self.chunks[self.current_chunk_idx], count as u16, self.line);
            self.emit_global_set_with_cls(&name);
        }

        Ok(())
    }

    fn compile_initial_value(&mut self, pic: &Option<String>, value: &Option<Literal>) -> Result<(), String> {
        if let Some(val) = value {
            self.compile_literal(val)?;
            // If PIC X and value is string, pad with spaces to PIC size
            if let (Some(pic), Literal::Str(s)) = (pic, val) {
                let size = Self::pic_size(pic);
                if size > 0 && s.len() < size {
                    // Already compiled the literal; now pad it
                    // Emit: str_pad_end to pad with spaces
                    self.emit_constant(Value::F64(size as f64));
                    self.emit_constant(Value::String(Rc::from(" ")));
                    self.emit(Op::str_pad_end);
                }
            }
        } else if let Some(pic) = pic {
            // Default based on PIC: X → spaces, 9 → 0
            let upper = pic.to_uppercase();
            let size = Self::pic_size(pic);
            if upper.starts_with('X') || upper.starts_with('A') {
                // Space-fill to PIC size
                let spaces: String = " ".repeat(size.max(1));
                self.emit_constant(Value::String(Rc::from(spaces.as_str())));
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
                    // For variable references, apply PIC editing if available
                    if let Expr::Ident(var_name) = expr {
                        let pic_key = format!("__PIC_{}", var_name);
                        let pk = self.add_string_constant(&pic_key);
                        self.emit_u16(Op::global_get, pk);
                        self.emit(Op::dup);
                        self.emit(Op::ref_is_null);
                        let no_pic = self.emit_jump(Op::br_if_true);
                        // Has PIC — format the value
                        let pic_slot = self.next_local; self.next_local += 1;
                        self.emit_u16(Op::local_set, pic_slot);
                        self.compile_expr(expr)?;
                        self.emit_u16(Op::local_get, pic_slot);
                        let fi = self.import("vybe:string", "format");
                        self.emit_host_call(fi, 2);
                        let skip = self.emit_jump(Op::br);
                        // No PIC — just compile normally
                        self.patch_jump(no_pic);
                        self.emit(Op::drop);
                        self.compile_expr(expr)?;
                        self.patch_jump(skip);
                    } else {
                        self.compile_expr(expr)?;
                    }
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
                    // Apply COBOL MOVE semantics: pad based on target PIC
                    self.emit_move_with_padding(dst)?;
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
                common_collections::emit_len(&mut self.chunks[self.current_chunk_idx], self.line);
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
                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                let key_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, key_slot);
                // dst[key] = src[key]
                self.emit_u16(Op::global_get, dst_idx);
                self.emit_u16(Op::global_get, src_idx);
                self.emit_u16(Op::local_get, key_slot);
                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line); // src[key]
                self.emit_u16(Op::local_get, key_slot);
                // struct_set expects [obj, val] with key as constant — use array_set instead
                common_collections::emit_set(&mut self.chunks[self.current_chunk_idx], self.line);
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
                { let c = self.current_chunk_idx; let line = self.line; common::math::emit_trunc(&mut self.chunks[c], line); }
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
                // Apply COMP-3 rounding if target has PIC with V (decimal)
                let pic_key = format!("__PIC_{}", dst);
                let pk = self.add_string_constant(&pic_key);
                self.emit_u16(Op::global_get, pk);
                self.emit(Op::dup);
                self.emit(Op::ref_is_null);
                let no_pic = self.emit_jump(Op::br_if_true);
                // Has PIC — check for V (implied decimal)
                // The PIC string is on stack, but we need it as a Rust string
                // to determine scale. Since we can't inspect at compile time
                // for dynamic vars, emit a generic round-to-2 for V99 patterns.
                self.emit(Op::drop); // drop PIC string
                // Round to 2 decimal places (most common: V99)
                self.emit_constant(Value::F64(100.0));
                self.emit(Op::f64_mul);
                let c = self.current_chunk_idx;
                let line = self.line;
                common::math::emit_round(&mut self.chunks[c], line);
                self.emit_constant(Value::F64(100.0));
                self.emit(Op::f64_div);
                let end = self.emit_jump(Op::br);
                self.patch_jump(no_pic);
                self.emit(Op::drop); // drop null
                self.patch_jump(end);
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

            Statement::PerformUntil { test, body, test_after } => {
                if *test_after {
                    // WITH TEST AFTER: do-while (execute body at least once)
                    let loop_start = self.current_offset();
                    for s in body { self.compile_statement(s)?; }
                    self.compile_expr(test)?;
                    self.emit(Op::dyn_to_bool);
                    self.emit(Op::dyn_not); // loop while NOT condition
                    let exit = self.emit_jump(Op::br_if_false);
                    self.emit_loop(loop_start);
                    self.patch_jump(exit);
                } else {
                    // WITH TEST BEFORE (default): test-then-loop
                    let loop_start = self.current_offset();
                    self.compile_expr(test)?;
                    self.emit(Op::dyn_to_bool);
                    let exit = self.emit_jump(Op::br_if_true);
                    for s in body { self.compile_statement(s)?; }
                    self.emit_loop(loop_start);
                    self.patch_jump(exit);
                }
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
                    // Paragraph not found locally — emit as import call
                    let i = self.import("*", name);
                    self.emit_host_call(i, 0);
                    self.emit(Op::drop);
                }
            }

            Statement::StringConcat { sources, into, pointer: _ } => {
                let mut first = true;
                for source in sources {
                    self.compile_expr(&source.value)?;
                    if !first { common_strings::emit_str_concat(&mut self.chunks[self.current_chunk_idx], self.line); }
                    first = false;
                }
                if first { self.emit_constant(Value::String(Rc::from(""))); }
                let idx = self.add_string_constant(into);
                self.emit_u16(Op::global_set, idx);
            }

            Statement::Unstring { src, delimiters, into, pointer: _ } => {
                let src_idx = self.add_string_constant(src);
                self.emit_u16(Op::global_get, src_idx);
                if let Some(delim) = delimiters.first() {
                    self.emit_constant(Value::String(Rc::from(delim.as_str())));
                } else {
                    self.emit_constant(Value::String(Rc::from(" ")));
                }
                common_strings::emit_split(&mut self.chunks[self.current_chunk_idx], self.line);
                // Assign each part to the target variables
                let arr_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, arr_slot);
                for (i, target) in into.iter().enumerate() {
                    self.emit_u16(Op::local_get, arr_slot);
                    self.emit_constant(Value::I32(i as i32));
                    common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                    let idx = self.add_string_constant(&target.name);
                    self.emit_u16(Op::global_set, idx);
                }
            }

            Statement::InspectTallying { var, counter, mode: _, target } => {
                // Count occurrences of target in var
                let var_idx = self.add_string_constant(var);
                self.emit_u16(Op::global_get, var_idx);
                self.emit_constant(Value::String(Rc::from(target.as_str())));
                common_strings::emit_split(&mut self.chunks[self.current_chunk_idx], self.line);
                common_collections::emit_len(&mut self.chunks[self.current_chunk_idx], self.line);
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
                common_strings::emit_replace(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit_u16(Op::global_set, var_idx);
            }

            Statement::Call { name, args } => {
                for arg in args {
                    let ai = self.add_string_constant(arg);
                    self.emit_u16(Op::global_get, ai);
                }
                // Check cross-language common imports first, then fall back to wildcard
                let i = if let Some((module, func)) = vybe_compiler_common::imports::resolve_common_import(name) {
                    self.import(module, func)
                } else {
                    self.import("*", name)
                };
                self.emit_host_call(i, args.len() as u8);
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
                    AcceptSource::CommandLine => {
                        let i = self.import("wasi:cli", "getArgs");
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
                common_strings::emit_replace(&mut self.chunks[self.current_chunk_idx], self.line);
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
                // CALL "program" ASYNC → call_import then spawn thread
                // Push args
                for arg in args {
                    let ai = self.add_string_constant(arg);
                    self.emit_u16(Op::global_get, ai);
                }
                let i = self.import("*", name);
                self.emit_host_call(i, args.len() as u8);
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
                // RUN UNIT → call external program via import
                for arg in args {
                    let ai = self.add_string_constant(arg);
                    self.emit_u16(Op::global_get, ai);
                }
                let i = self.import("*", name);
                self.emit_host_call(i, args.len() as u8);
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
                self.emit(Op::null);
                self.emit_u16(Op::suspend, 0);
            }

            Statement::AddCorresponding { src, dst } => {
                // ADD CORRESPONDING src TO dst → add matching fields
                // Same pattern as MOVE CORRESPONDING but with addition
                let src_idx = self.add_string_constant(src);
                let dst_idx = self.add_string_constant(dst);
                self.emit_u16(Op::global_get, dst_idx);
                self.emit_u16(Op::global_get, src_idx);
                self.emit(Op::dyn_add);
                self.emit_u16(Op::global_set, dst_idx);
            }

            Statement::SubtractCorresponding { src, dst } => {
                let dst_idx = self.add_string_constant(dst);
                let src_idx = self.add_string_constant(src);
                self.emit_u16(Op::global_get, dst_idx);
                self.emit_u16(Op::global_get, src_idx);
                self.emit(Op::f64_sub);
                self.emit_u16(Op::global_set, dst_idx);
            }

            Statement::CopyReplacing { copybook: _, replacements: _ } => {
                // COPY REPLACING is a preprocessor directive — no runtime effect
            }

            Statement::AcceptCommandLine(var) => {
                let i = self.import("wasi:cli", "getArgs");
                self.emit_host_call(i, 0);
                let idx = self.add_string_constant(var);
                self.emit_u16(Op::global_set, idx);
            }

            // ── CICS ───────────────────────────────────────────
            Statement::CicsCommand { command, params } => {
                // Map CICS commands to host functions
                match command.as_str() {
                    "SEND" => {
                        // SEND MAP(mapname) MAPSET(setname) — display screen
                        for (key, val) in params {
                            if key == "FROM" || key == "MAP" {
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_get, vi);
                                let c = self.current_chunk_idx;
                                let line = self.line;
                                common::io::emit_print(&mut self.chunks[c], 1, line);
                                self.emit(Op::drop);
                            }
                        }
                    }
                    "RECEIVE" => {
                        // RECEIVE MAP(mapname) INTO(dataname) — read screen
                        for (key, val) in params {
                            if key == "INTO" {
                                let i = self.import("wasi:cli", "readLine");
                                self.emit_host_call(i, 0);
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_set, vi);
                            }
                        }
                    }
                    "READ" => {
                        // READ FILE(filename) INTO(dataname) RIDFLD(key)
                        for (key, val) in params {
                            if key == "INTO" {
                                let i = self.import("wasi:cli", "readLine");
                                self.emit_host_call(i, 0);
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_set, vi);
                            }
                        }
                    }
                    "WRITE" => {
                        for (key, val) in params {
                            if key == "FROM" {
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_get, vi);
                                let c = self.current_chunk_idx;
                                let line = self.line;
                                common::io::emit_print(&mut self.chunks[c], 1, line);
                                self.emit(Op::drop);
                            }
                        }
                    }
                    "RETURN" => {
                        // RETURN TRANSID(next-trans) — return to CICS
                        self.emit(Op::null);
                        self.emit(Op::r#return);
                    }
                    "LINK" | "XCTL" => {
                        // LINK/XCTL PROGRAM(progname) — call another program
                        for (key, val) in params {
                            if key == "PROGRAM" {
                                let ni = self.add_string_constant(val);
                                self.emit_u16(Op::global_get, ni);
                                self.emit_u8(Op::call_ref, 0);
                                self.emit(Op::drop);
                            }
                        }
                    }
                    "STARTBR" | "READNEXT" | "READPREV" | "ENDBR" | "REWRITE" | "DELETE" | "UNLOCK" => {
                        // File browsing — simplified no-ops
                    }
                    "ASKTIME" | "FORMATTIME" => {
                        // Time functions
                        for (key, val) in params {
                            if key == "ABSTIME" || key == "DDMMYYYY" || key == "TIME" {
                                let i = self.import("wasi:clocks", "toISOString");
                                self.emit_host_call(i, 0);
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_set, vi);
                            }
                        }
                    }
                    "GETMAIN" => {
                        // Allocate memory — create empty object
                        for (key, val) in params {
                            if key == "SET" {
                                let line = self.line;
                                let c = self.current_chunk_idx;
                                common::dict::emit_new(&mut self.chunks[c], line);
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_set, vi);
                            }
                        }
                    }
                    "FREEMAIN" => {
                        // Free memory — set to null
                        for (key, val) in params {
                            if key == "DATA" || key == "DATAPOINTER" {
                                self.emit(Op::null);
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_set, vi);
                            }
                        }
                    }

                    // ── HANDLE CONDITION (error routing) ───────
                    "HANDLE" => {
                        // HANDLE CONDITION ERROR(para) NOTFND(para) etc.
                        // Store handler paragraph names as globals for later dispatch
                        for (condition, handler) in params {
                            let key = format!("__CICS_HANDLER_{}", condition);
                            let ki = self.add_string_constant(&key);
                            self.emit_constant(Value::String(Rc::from(handler.as_str())));
                            self.emit_u16(Op::global_set, ki);
                        }
                    }

                    // ── HANDLE AID (key handling) ──────────────
                    // HANDLE AID PF1(para) PF3(para) ENTER(para) CLEAR(para)
                    // Same pattern — store handler names
                    // (handled by "HANDLE" above since lexer puts it all together)

                    // ── WRITEQ TS (temporary storage queue) ────
                    "WRITEQ" => {
                        // WRITEQ TS QUEUE(name) FROM(data) [ITEM(n)]
                        let mut queue_name = String::new();
                        let mut from_var = String::new();
                        for (key, val) in params {
                            match key.as_str() {
                                "TS" | "TD" => {} // just the queue type indicator
                                "QUEUE" => queue_name = val.clone(),
                                "FROM" => from_var = val.clone(),
                                _ => {}
                            }
                        }
                        if !queue_name.is_empty() {
                            // Queue is stored as a global array
                            let qk = self.add_string_constant(&format!("__CICS_Q_{}", queue_name));
                            // Get or create the queue array
                            self.emit_u16(Op::global_get, qk);
                            self.emit(Op::dup);
                            self.emit(Op::ref_is_null);
                            let exists = self.emit_jump(Op::br_if_false);
                            self.emit(Op::drop);
                            common_collections::emit_array_new(&mut self.chunks[self.current_chunk_idx], 0, self.line);
                            self.patch_jump(exists);
                            // Push data to queue
                            if !from_var.is_empty() {
                                let fi = self.add_string_constant(&from_var);
                                self.emit_u16(Op::global_get, fi);
                            } else {
                                self.emit_constant(Value::String(Rc::from("")));
                            }
                            common_collections::emit_push(&mut self.chunks[self.current_chunk_idx], self.line);
                            self.emit(Op::drop);
                            // Save queue back
                            self.emit_u16(Op::global_set, qk);
                        }
                    }

                    // ── READQ TS (read from temp storage queue) ─
                    "READQ" => {
                        // READQ TS QUEUE(name) INTO(data) [ITEM(n)]
                        let mut queue_name = String::new();
                        let mut into_var = String::new();
                        let mut item_num: Option<String> = None;
                        for (key, val) in params {
                            match key.as_str() {
                                "TS" | "TD" => {}
                                "QUEUE" => queue_name = val.clone(),
                                "INTO" => into_var = val.clone(),
                                "ITEM" => item_num = Some(val.clone()),
                                _ => {}
                            }
                        }
                        if !queue_name.is_empty() {
                            let qk = self.add_string_constant(&format!("__CICS_Q_{}", queue_name));
                            self.emit_u16(Op::global_get, qk);
                            // Get item by index (default: first item = shift)
                            if let Some(item) = &item_num {
                                let ii = self.add_string_constant(item);
                                self.emit_u16(Op::global_get, ii);
                                self.emit_constant(Value::I32(1));
                                self.emit(Op::f64_sub); // CICS is 1-indexed
                                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                            } else {
                                // Read next = shift from front
                                let c = self.current_chunk_idx;
                                let line = self.line;
                                common::collections::emit_shift(&mut self.chunks[c], line);
                            }
                            if !into_var.is_empty() {
                                let vi = self.add_string_constant(&into_var);
                                self.emit_u16(Op::global_set, vi);
                            } else {
                                self.emit(Op::drop);
                            }
                        }
                    }

                    // ── DELETEQ TS (delete temp storage queue) ──
                    "DELETEQ" => {
                        for (key, val) in params {
                            if key == "QUEUE" {
                                let qk = self.add_string_constant(&format!("__CICS_Q_{}", val));
                                self.emit(Op::null);
                                self.emit_u16(Op::global_set, qk);
                            }
                        }
                    }

                    // ── ENQ / DEQ (named resource locking) ─────
                    "ENQ" => {
                        // ENQ RESOURCE(name) LENGTH(n)
                        for (key, val) in params {
                            if key == "RESOURCE" {
                                let rk = self.add_string_constant(&format!("__CICS_LOCK_{}", val));
                                // Simple spinlock via global flag
                                self.emit_constant(Value::I32(1));
                                self.emit_u16(Op::global_set, rk);
                            }
                        }
                    }
                    "DEQ" => {
                        for (key, val) in params {
                            if key == "RESOURCE" {
                                let rk = self.add_string_constant(&format!("__CICS_LOCK_{}", val));
                                self.emit_constant(Value::I32(0));
                                self.emit_u16(Op::global_set, rk);
                            }
                        }
                    }

                    // ── ASSIGN (system info) ───────────────────
                    "ASSIGN" => {
                        // ASSIGN USERID(var) TERMINAL(var) SYSID(var) etc.
                        for (key, val) in params {
                            match key.as_str() {
                                "USERID" => {
                                    let i = self.import("wasi:cli", "userName");
                                    self.emit_host_call(i, 0);
                                    let vi = self.add_string_constant(val);
                                    self.emit_u16(Op::global_set, vi);
                                }
                                "SYSID" | "APPLID" => {
                                    self.emit_constant(Value::String(Rc::from("VYBE")));
                                    let vi = self.add_string_constant(val);
                                    self.emit_u16(Op::global_set, vi);
                                }
                                "CWALENGTH" => {
                                    self.emit_constant(Value::F64(0.0));
                                    let vi = self.add_string_constant(val);
                                    self.emit_u16(Op::global_set, vi);
                                }
                                _ => {
                                    // Generic: return empty string
                                    self.emit_constant(Value::String(Rc::from("")));
                                    let vi = self.add_string_constant(val);
                                    self.emit_u16(Op::global_set, vi);
                                }
                            }
                        }
                    }

                    // ── COMMAREA (communication area) ──────────
                    // COMMAREA is passed via LINK/XCTL — it's just data
                    // Our implementation uses global variables as the shared area.
                    // LINK PROGRAM(X) COMMAREA(data) LENGTH(n) → call with data as arg

                    // ── PUT/GET CONTAINER (Channels) ───────────
                    "PUT" => {
                        // PUT CONTAINER(name) CHANNEL(ch) FROM(data)
                        let mut container = String::new();
                        let mut channel = String::new();
                        let mut from_var = String::new();
                        for (key, val) in params {
                            match key.as_str() {
                                "CONTAINER" => container = val.clone(),
                                "CHANNEL" => channel = val.clone(),
                                "FROM" => from_var = val.clone(),
                                _ => {}
                            }
                        }
                        if !container.is_empty() {
                            let ck = self.add_string_constant(&format!("__CICS_CONT_{}_{}", channel, container));
                            if !from_var.is_empty() {
                                let fi = self.add_string_constant(&from_var);
                                self.emit_u16(Op::global_get, fi);
                            } else {
                                self.emit_constant(Value::String(Rc::from("")));
                            }
                            self.emit_u16(Op::global_set, ck);
                        }
                    }
                    "GET" => {
                        // GET CONTAINER(name) CHANNEL(ch) INTO(data)
                        let mut container = String::new();
                        let mut channel = String::new();
                        let mut into_var = String::new();
                        for (key, val) in params {
                            match key.as_str() {
                                "CONTAINER" => container = val.clone(),
                                "CHANNEL" => channel = val.clone(),
                                "INTO" | "SET" => into_var = val.clone(),
                                _ => {}
                            }
                        }
                        if !container.is_empty() {
                            let ck = self.add_string_constant(&format!("__CICS_CONT_{}_{}", channel, container));
                            self.emit_u16(Op::global_get, ck);
                            if !into_var.is_empty() {
                                let vi = self.add_string_constant(&into_var);
                                self.emit_u16(Op::global_set, vi);
                            } else {
                                self.emit(Op::drop);
                            }
                        }
                    }

                    // ── DELAY / START (time control) ───────────
                    "DELAY" => {
                        // DELAY INTERVAL(hhmmss) or DELAY FOR SECONDS(n)
                        for (key, val) in params {
                            if key == "SECONDS" || key == "INTERVAL" {
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_get, vi);
                                let i = self.import("wasi:clocks", "sleep");
                                self.emit_host_call(i, 1);
                            }
                        }
                    }
                    "START" => {
                        // START TRANSID(txn) — schedule transaction
                        for (key, val) in params {
                            if key == "TRANSID" {
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_get, vi);
                                self.emit_u8(Op::call_ref, 0);
                                self.emit(Op::drop);
                            }
                        }
                    }

                    // ── SUSPEND / POST / WAIT EVENT ────────────
                    "SUSPEND" => {
                        self.emit(Op::null);
                        self.emit_u16(Op::suspend, 0);
                    }
                    "POST" => {
                        // POST EVENT(name) — signal event
                        for (key, val) in params {
                            if key == "EVENT" {
                                let ek = self.add_string_constant(&format!("__CICS_EVT_{}", val));
                                self.emit(Op::r#true);
                                self.emit_u16(Op::global_set, ek);
                            }
                        }
                    }
                    "WAIT" => {
                        // WAIT EVENT — wait for posted event
                        // Simplified: just continue (real impl would suspend)
                    }

                    // ── WEB (CICS Web Services) ────────────────
                    "WEB" => {
                        // WEB SEND/RECEIVE — HTTP operations
                        for (key, val) in params {
                            if key == "FROM" {
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_get, vi);
                                let c = self.current_chunk_idx;
                                let line = self.line;
                                common::io::emit_print(&mut self.chunks[c], 1, line);
                                self.emit(Op::drop);
                            }
                            if key == "INTO" || key == "SET" {
                                let i = self.import("wasi:cli", "readLine");
                                self.emit_host_call(i, 0);
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_set, vi);
                            }
                        }
                    }

                    // ── DOCUMENT (CICS Document) ───────────────
                    "DOCUMENT" => {
                        // DOCUMENT CREATE/SET/INSERT — build response document
                        for (key, val) in params {
                            if key == "DOCTOKEN" || key == "SET" {
                                let line = self.line;
                                let c = self.current_chunk_idx;
                                common::dict::emit_new(&mut self.chunks[c], line);
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_set, vi);
                            }
                        }
                    }

                    // ── ABEND ──────────────────────────────────
                    "ABEND" => {
                        // ABEND ABCODE(code) — abnormal end
                        let mut code = "ASRA".to_string();
                        for (key, val) in params {
                            if key == "ABCODE" { code = val.clone(); }
                        }
                        self.emit_constant(Value::String(Rc::from(code.as_str())));
                        let line = self.line;
                        common::errors::emit_throw(&mut self.chunks[self.current_chunk_idx], line);
                    }

                    // ── INQUIRE / SET (system queries) ─────────
                    "INQUIRE" | "SET" => {
                        // System resource queries — simplified
                        for (key, val) in params {
                            if !val.is_empty() {
                                self.emit_constant(Value::String(Rc::from("")));
                                let vi = self.add_string_constant(val);
                                self.emit_u16(Op::global_set, vi);
                            }
                        }
                    }

                    _ => {
                        // Unknown CICS command — no-op
                    }
                }
            }

            Statement::DliCommand { command, params } => {
                // IMS/DLI commands
                match command.as_str() {
                    "GU" | "GN" | "GNP" | "GHU" | "GHN" => {
                        // Get Unique / Get Next — database read
                        for (key, val) in params {
                            if key == "INTO" || key.is_empty() {
                                let i = self.import("wasi:cli", "readLine");
                                self.emit_host_call(i, 0);
                                if !val.is_empty() {
                                    let vi = self.add_string_constant(val);
                                    self.emit_u16(Op::global_set, vi);
                                } else {
                                    self.emit(Op::drop);
                                }
                            }
                        }
                    }
                    "ISRT" | "REPL" | "DLET" => {
                        // Insert / Replace / Delete
                        for (key, val) in params {
                            if key == "FROM" || key.is_empty() {
                                if !val.is_empty() {
                                    let vi = self.add_string_constant(val);
                                    self.emit_u16(Op::global_get, vi);
                                    let c = self.current_chunk_idx;
                                    let line = self.line;
                                    common::io::emit_print(&mut self.chunks[c], 1, line);
                                    self.emit(Op::drop);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }

            // ── Arithmetic with clauses ────────────────────────
            Statement::AddRounded { srcs, to, giving } => {
                let to_idx = self.add_string_constant(to);
                if let Some(giving_name) = giving {
                    let mut first = true;
                    for src in srcs { self.compile_expr(src)?; if !first { self.emit(Op::dyn_add); } first = false; }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::math::emit_round(&mut self.chunks[c], line);
                    let idx = self.add_string_constant(giving_name);
                    self.emit_u16(Op::global_set, idx);
                } else {
                    self.emit_u16(Op::global_get, to_idx);
                    for src in srcs { self.compile_expr(src)?; self.emit(Op::dyn_add); }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::math::emit_round(&mut self.chunks[c], line);
                    self.emit_u16(Op::global_set, to_idx);
                }
            }

            Statement::ComputeWithError { dst, expr, rounded, on_error, not_on_error } => {
                // Try the computation, catch overflow
                let line = self.line;
                let c = self.current_chunk_idx;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[c], line);
                self.compile_expr(expr)?;
                if *rounded {
                    common::math::emit_round(&mut self.chunks[c], line);
                }
                let di = self.add_string_constant(dst);
                self.emit_u16(Op::global_set, di);
                let line = self.line;
                common::errors::emit_try_end(&mut self.chunks[c], line);
                // NOT ON SIZE ERROR
                for s in not_on_error { self.compile_statement(s)?; }
                let skip = self.emit_jump(Op::br);
                common::errors::patch_catch(&mut self.chunks[c], catch_jump);
                self.emit(Op::drop);
                // ON SIZE ERROR
                for s in on_error { self.compile_statement(s)?; }
                self.patch_jump(skip);
            }

            Statement::ReadFileAtEnd { file, into, at_end, not_at_end } => {
                let fi = self.add_string_constant(&format!("__file_{}", file));
                self.emit_u16(Op::global_get, fi);
                let i = self.import("wasi:filesystem", "lineInput");
                self.emit_host_call(i, 1);
                // Check if null (end of file)
                self.emit(Op::dup);
                self.emit(Op::ref_is_null);
                let eof_jump = self.emit_jump(Op::br_if_true);
                // NOT AT END — got data
                if let Some(var) = into {
                    let idx = self.add_string_constant(var);
                    self.emit_u16(Op::global_set, idx);
                } else {
                    self.emit(Op::drop);
                }
                for s in not_at_end { self.compile_statement(s)?; }
                let end = self.emit_jump(Op::br);
                // AT END
                self.patch_jump(eof_jump);
                self.emit(Op::drop);
                for s in at_end { self.compile_statement(s)?; }
                self.patch_jump(end);
            }

            Statement::NestedProgram(inner_prog) => {
                // Compile nested program as a separate set of chunks
                // Simplified: compile its main body inline
                for s in &inner_prog.main_body {
                    self.compile_statement(s)?;
                }
            }

            Statement::DisplayFormatted { var, pic } => {
                // Format a number using PIC editing mask
                let vi = self.add_string_constant(var);
                self.emit_u16(Op::global_get, vi);
                self.emit_constant(Value::String(Rc::from(pic.as_str())));
                let i = self.import("vybe:string", "format");
                self.emit_host_call(i, 2);
                let c = self.current_chunk_idx;
                let line = self.line;
                common::io::emit_print(&mut self.chunks[c], 1, line);
                self.emit(Op::drop);
            }

            // ── Embedded SQL ───────────────────────────────────
            Statement::SqlConnect { dsn, handle_var } => {
                // EXEC SQL CONNECT :dsn END-EXEC → vybe:database connect
                let di = self.add_string_constant(dsn);
                self.emit_u16(Op::global_get, di);
                let i = self.import("vybe:database", "connect");
                self.emit_host_call(i, 1);
                if let Some(hv) = handle_var {
                    let hi = self.add_string_constant(hv);
                    self.emit_u16(Op::global_set, hi);
                } else {
                    // Store as default connection
                    let idx = self.add_string_constant("__SQL_CONN");
                    self.emit_u16(Op::global_set, idx);
                }
            }

            Statement::SqlSelect { sql, into_vars, from_vars } => {
                // EXEC SQL SELECT ... INTO :var1, :var2 FROM ... WHERE :var3 END-EXEC
                // Build SQL string, push host vars, call query
                let conn_idx = self.add_string_constant("__SQL_CONN");
                self.emit_u16(Op::global_get, conn_idx);
                self.emit_constant(Value::String(Rc::from(sql.as_str())));
                // Push WHERE-clause host vars as parameters
                for var in from_vars {
                    let vi = self.add_string_constant(var);
                    self.emit_u16(Op::global_get, vi);
                }
                let i = self.import("vybe:database", "query");
                self.emit_host_call(i, (from_vars.len() + 2) as u8);
                // Result is array of rows. For SELECT INTO, take first row.
                let result_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, result_slot);
                // Assign each INTO var from the result
                for (col, var) in into_vars.iter().enumerate() {
                    self.emit_u16(Op::local_get, result_slot);
                    // First row: result[0]
                    self.emit_constant(Value::I32(0));
                    common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                    // Column: row[col]
                    self.emit_constant(Value::I32(col as i32));
                    common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                    let vi = self.add_string_constant(var);
                    self.emit_u16(Op::global_set, vi);
                }
                // Set SQLCODE = 0 (success)
                self.emit_constant(Value::F64(0.0));
                let sc = self.add_string_constant("SQLCODE");
                self.emit_u16(Op::global_set, sc);
            }

            Statement::SqlExecute { sql, host_vars } => {
                // EXEC SQL INSERT/UPDATE/DELETE ... END-EXEC
                let conn_idx = self.add_string_constant("__SQL_CONN");
                self.emit_u16(Op::global_get, conn_idx);
                self.emit_constant(Value::String(Rc::from(sql.as_str())));
                for var in host_vars {
                    let vi = self.add_string_constant(var);
                    self.emit_u16(Op::global_get, vi);
                }
                let i = self.import("vybe:database", "execute");
                self.emit_host_call(i, (host_vars.len() + 2) as u8);
                self.emit(Op::drop);
                // SQLCODE = 0
                self.emit_constant(Value::F64(0.0));
                let sc = self.add_string_constant("SQLCODE");
                self.emit_u16(Op::global_set, sc);
            }

            Statement::SqlCommit => {
                let conn_idx = self.add_string_constant("__SQL_CONN");
                self.emit_u16(Op::global_get, conn_idx);
                self.emit_constant(Value::String(Rc::from("COMMIT")));
                let i = self.import("vybe:database", "execute");
                self.emit_host_call(i, 2);
                self.emit(Op::drop);
            }

            Statement::SqlRollback => {
                let conn_idx = self.add_string_constant("__SQL_CONN");
                self.emit_u16(Op::global_get, conn_idx);
                self.emit_constant(Value::String(Rc::from("ROLLBACK")));
                let i = self.import("vybe:database", "execute");
                self.emit_host_call(i, 2);
                self.emit(Op::drop);
            }

            Statement::SqlDeclareCursor { cursor_name, sql, host_vars } => {
                // Store cursor SQL + params for later OPEN
                let cursor_sql_key = self.add_string_constant(&format!("__CURSOR_{}_SQL", cursor_name));
                self.emit_constant(Value::String(Rc::from(sql.as_str())));
                self.emit_u16(Op::global_set, cursor_sql_key);
                // Store host var names for parameter binding
                for (i, var) in host_vars.iter().enumerate() {
                    let pk = self.add_string_constant(&format!("__CURSOR_{}_P{}", cursor_name, i));
                    let vi = self.add_string_constant(var);
                    self.emit_u16(Op::global_get, vi);
                    self.emit_u16(Op::global_set, pk);
                }
                let pc = self.add_string_constant(&format!("__CURSOR_{}_PCNT", cursor_name));
                self.emit_constant(Value::F64(host_vars.len() as f64));
                self.emit_u16(Op::global_set, pc);
            }

            Statement::SqlOpenCursor(cursor_name) => {
                // Execute the cursor's SQL and store result set
                let conn_idx = self.add_string_constant("__SQL_CONN");
                self.emit_u16(Op::global_get, conn_idx);
                let sql_key = self.add_string_constant(&format!("__CURSOR_{}_SQL", cursor_name));
                self.emit_u16(Op::global_get, sql_key);
                let i = self.import("vybe:database", "query");
                self.emit_host_call(i, 2);
                // Store result set
                let rs_key = self.add_string_constant(&format!("__CURSOR_{}_RS", cursor_name));
                self.emit_u16(Op::global_set, rs_key);
                // Reset row index to 0
                let idx_key = self.add_string_constant(&format!("__CURSOR_{}_IDX", cursor_name));
                self.emit_constant(Value::F64(0.0));
                self.emit_u16(Op::global_set, idx_key);
            }

            Statement::SqlFetch { cursor_name, into_vars } => {
                // Fetch next row from cursor result set
                let rs_key = self.add_string_constant(&format!("__CURSOR_{}_RS", cursor_name));
                let idx_key = self.add_string_constant(&format!("__CURSOR_{}_IDX", cursor_name));

                // Check if index < result set length
                self.emit_u16(Op::global_get, idx_key);
                self.emit_u16(Op::global_get, rs_key);
                common_collections::emit_len(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit(Op::dyn_lt);
                let no_more = self.emit_jump(Op::br_if_false);

                // Get current row
                self.emit_u16(Op::global_get, rs_key);
                self.emit_u16(Op::global_get, idx_key);
                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                let row_slot = self.next_local; self.next_local += 1;
                self.emit_u16(Op::local_set, row_slot);

                // Assign each column to INTO var
                for (col, var) in into_vars.iter().enumerate() {
                    self.emit_u16(Op::local_get, row_slot);
                    self.emit_constant(Value::I32(col as i32));
                    common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                    let vi = self.add_string_constant(var);
                    self.emit_u16(Op::global_set, vi);
                }

                // Increment index
                self.emit_u16(Op::global_get, idx_key);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::global_set, idx_key);

                // SQLCODE = 0 (success)
                self.emit_constant(Value::F64(0.0));
                let sc = self.add_string_constant("SQLCODE");
                self.emit_u16(Op::global_set, sc);
                let end = self.emit_jump(Op::br);

                // No more rows: SQLCODE = 100
                self.patch_jump(no_more);
                self.emit_constant(Value::F64(100.0));
                let sc2 = self.add_string_constant("SQLCODE");
                self.emit_u16(Op::global_set, sc2);

                self.patch_jump(end);
            }

            Statement::SqlCloseCursor(cursor_name) => {
                // Clear cursor state
                let rs_key = self.add_string_constant(&format!("__CURSOR_{}_RS", cursor_name));
                self.emit(Op::null);
                self.emit_u16(Op::global_set, rs_key);
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
                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
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
                common_strings::emit_substring(&mut self.chunks[self.current_chunk_idx], self.line);
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
                        common_strings::emit_length(&mut self.chunks[self.current_chunk_idx], self.line);
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

    // ------------------------------------------------------------------
    // COBOL MOVE semantics — space-padding and zero-filling
    // ------------------------------------------------------------------

    /// MOVE with COBOL padding semantics:
    /// - Alpha (PIC X/A): right-fill with spaces, truncate on right
    /// - Numeric (PIC 9): left-fill with zeros, truncate on left
    /// - Alphanumeric moves: convert to string, pad
    /// Stack: [value] → [] (stores in global)
    fn emit_move_with_padding(&mut self, dst: &str) -> Result<(), String> {
        // Check if we have PIC metadata for the target
        let pic_key = format!("__PIC_{}", dst);
        let pk = self.add_string_constant(&pic_key);
        self.emit_u16(Op::global_get, pk);
        self.emit(Op::dup);
        self.emit(Op::ref_is_null);
        let has_pic = self.emit_jump(Op::br_if_false);

        // No PIC metadata — just store raw value
        self.emit(Op::drop); // drop null PIC
        self.emit_global_set_with_cls(dst);
        let end = self.emit_jump(Op::br);

        // Has PIC metadata — apply padding
        self.patch_jump(has_pic);
        let pic_slot = self.next_local; self.next_local += 1;
        self.emit_u16(Op::local_set, pic_slot); // PIC string
        // Stack: [value]
        let val_slot = self.next_local; self.next_local += 1;
        self.emit_u16(Op::local_set, val_slot);

        // Parse PIC to determine type and size
        // Call the PIC padding helper
        self.emit_u16(Op::local_get, val_slot);
        self.emit_u16(Op::local_get, pic_slot);
        let i = self.import("vybe:string", "format");
        self.emit_host_call(i, 2);
        self.emit_global_set_with_cls(dst);

        self.patch_jump(end);
        Ok(())
    }

    /// Emit PIC editing format for DISPLAY.
    /// Takes a raw value and PIC string on stack, produces formatted string.
    /// Handles: Z (zero suppress), $ (currency), , (comma), . (decimal), 9 (digit), - (sign)
    /// Stack: [value, pic_string] → [formatted_string]
    fn emit_pic_format(&mut self, _pic: &str) {
        // Use host string format function which handles basic formatting
        let i = self.import("vybe:string", "format");
        self.emit_host_call(i, 2);
    }

    // ------------------------------------------------------------------
    // COMP-3 packed decimal support
    // ------------------------------------------------------------------
    // COMP-3 (packed decimal) stores digits in BCD format.
    // In our VM, all numbers are f64. For COMP-3 semantics,
    // we ensure arithmetic rounds to the PIC precision after each operation.
    // This is done by emitting a round-to-scale operation after arithmetic
    // when the target has a V (implied decimal) in its PIC.

    /// Extract size from PIC string. E.g. PIC X(20) → 20, PIC 9(5)V99 → 7
    fn pic_size(pic: &str) -> usize {
        let upper = pic.to_uppercase();
        let mut size = 0usize;
        let chars: Vec<char> = upper.chars().collect();
        let mut i = 0;
        while i < chars.len() {
            let c = chars[i];
            match c {
                'X' | '9' | 'A' | 'Z' | '$' | '-' | '+' | '.' | ',' | 'B' | '0' | '/' | '*' => {
                    // Check for repeat count: X(20)
                    if i + 1 < chars.len() && chars[i + 1] == '(' {
                        let mut num = String::new();
                        i += 2;
                        while i < chars.len() && chars[i] != ')' {
                            num.push(chars[i]);
                            i += 1;
                        }
                        if i < chars.len() { i += 1; } // skip )
                        size += num.parse::<usize>().unwrap_or(1);
                    } else {
                        size += 1;
                        i += 1;
                    }
                }
                'V' => { i += 1; } // implied decimal — no display character
                'S' => { i += 1; } // sign — no display character (unless SIGN LEADING SEPARATE)
                _ => { i += 1; }
            }
        }
        size
    }

    /// After arithmetic, round to COMP-3 precision based on PIC.
    /// PIC 9(5)V99 → round to 2 decimal places.
    fn emit_comp3_round(&mut self, pic: &str) {
        // Count digits after V to determine scale
        let upper = pic.to_uppercase();
        if let Some(v_pos) = upper.find('V') {
            let after_v = &upper[v_pos + 1..];
            let scale = after_v.chars().filter(|c| *c == '9').count();
            if scale > 0 {
                // Multiply by 10^scale, truncate, divide by 10^scale
                let factor = 10f64.powi(scale as i32);
                self.emit_constant(Value::F64(factor));
                self.emit(Op::f64_mul);
                let c = self.current_chunk_idx;
                let line = self.line;
                common::math::emit_round(&mut self.chunks[c], line);
                self.emit_constant(Value::F64(factor));
                self.emit(Op::f64_div);
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
            "LENGTH" => { common_strings::emit_length(&mut self.chunks[c], line); }
            "UPPER-CASE" => { common_strings::emit_to_upper(&mut self.chunks[c], line); }
            "LOWER-CASE" => { common_strings::emit_to_lower(&mut self.chunks[c], line); }
            "TRIM" => { common_strings::emit_trim(&mut self.chunks[c], line); }
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
                common_strings::emit_replace(&mut self.chunks[c], line);
            }
            "SQRT" => { common::math::emit_sqrt(&mut self.chunks[c], line); }
            "ABS" => { common::math::emit_abs(&mut self.chunks[c], line); }
            "SUM" => {
                // SUM — args on stack, add them up
                for _ in 1..args.len() { self.emit(Op::dyn_add); }
            }
            "INTEGER" => {
                let c = self.current_chunk_idx;
                let line = self.line;
                common::math::emit_trunc(&mut self.chunks[c], line);
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
                for _ in 1..args.len() { common_strings::emit_str_concat(&mut self.chunks[c], line); }
            }
            "TEST-NUMVAL" => {
                common::convert::emit_is_numeric(&mut self.chunks[c], line);
            }
            // Financial
            "ANNUITY" | "PRESENT-VALUE" => {
                for _ in 1..args.len() { self.emit(Op::drop); }
            }
            _ => {
                // Unknown function — call as import
                // args already on stack from the loop above
                let i = self.import("*", name);
                self.emit_host_call(i, args.len() as u8);
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

        // Stamp __types array for instanceof support
        common::classes::emit_instanceof_chain(&mut self.chunks[c], this_idx, class_name, line);

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
