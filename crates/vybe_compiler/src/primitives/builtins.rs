//! Builtin / opcode / intrinsic call compilation.
//!
//! Extracted from `primitives/mod.rs` (one `impl Compiler` block) to keep that
//! file navigable — same pattern as `calls.rs`/`classes.rs`. Methods are
//! private-by-convention, called from the core compile paths in `mod.rs`.

use super::*;

/// Compile-time `u8` from a literal instruction argument (a SIMD lane index).
fn expr_const_u8(expr: Option<&Expression>) -> u8 {
    match expr.map(|e| &e.kind) {
        Some(ExprKind::Lit(Literal::Int(n))) => *n as u8,
        _ => 0,
    }
}

/// Compile-time `u16` from a literal instruction argument (a GC type index).
/// A symbolic type name (`$T`) yields 0 — the VM's array ops carry the typeidx
/// only for binary-format parity and don't dispatch on it.
fn expr_const_u16(expr: Option<&Expression>) -> u16 {
    match expr.map(|e| &e.kind) {
        Some(ExprKind::Lit(Literal::Int(n))) => *n as u16,
        _ => 0,
    }
}

/// The textual form of a WASM type reference operand — a symbolic id (`$t` →
/// `"t"`) or a numeric type index (`3` → `"3"`). Used to key GC array-type
/// registration so each declared `(array …)` type maps to one registry id.
fn wasm_type_ref_name(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Ident(n) => n.clone(),
        ExprKind::Lit(Literal::Int(i)) => i.to_string(),
        _ => String::new(),
    }
}

/// A string literal argument's value (empty when not a string literal). Used
/// for the wast struct-type registration directive's name/parent fields.
fn expr_str_lit(expr: Option<&Expression>) -> String {
    match expr.map(|e| &e.kind) {
        Some(ExprKind::Lit(Literal::Str(s))) => s.to_string(),
        _ => String::new(),
    }
}

/// The named heap type a `ref.test`/`ref.cast` operand refers to, plus whether
/// it is nullable. A bare `$T` arrives as `Ident("T")` → `(name, false)`; a
/// folded `(ref null $T)` carries a `null` marker → `(name, true)`. The name is
/// recovered from the first ident found so either shape resolves.
fn wasm_heap_type_ref(expr: Option<&Expression>) -> (String, bool) {
    let Some(e) = expr else {
        return (String::new(), false);
    };
    match &e.kind {
        ExprKind::Ident(n) => (n.clone(), false),
        ExprKind::Lit(Literal::Str(s)) => (s.to_string(), false),
        // A folded ref type lowers to a call like `ref(null, $T)` / an object
        // carrying the heap type; dig for the type name and a `null` nullability
        // marker among the sub-expressions.
        _ => {
            let mut name = String::new();
            let mut nullable = false;
            collect_heap_type_ref(e, &mut name, &mut nullable);
            (name, nullable)
        }
    }
}

/// Walk a folded ref-type expression collecting the first non-`null` ident as
/// the type name and noting whether a `null` keyword appears (nullable).
fn collect_heap_type_ref(e: &Expression, name: &mut String, nullable: &mut bool) {
    match &e.kind {
        ExprKind::Ident(n) if n == "null" => *nullable = true,
        // `ref` is the reftype-constructor keyword in a folded `(ref [null] ht)`
        // operand, NOT a heap-type name — skip it so the real heap type (`i31`,
        // `$T`, …) is what gets recorded.
        ExprKind::Ident(n) if n == "ref" => {}
        ExprKind::Ident(n) if name.is_empty() => *name = n.clone(),
        ExprKind::Lit(Literal::Null) => *nullable = true,
        ExprKind::Lit(Literal::Str(s)) if &**s == "null" => *nullable = true,
        ExprKind::Lit(Literal::Str(s)) if name.is_empty() => *name = s.to_string(),
        ExprKind::Call { callee, args, .. } => {
            collect_heap_type_ref(callee, name, nullable);
            for a in args {
                collect_heap_type_ref(&a.value, name, nullable);
            }
        }
        _ => {}
    }
}

/// Encode a `v128.const` immediate: a shape token (`i32x4`, `f64x2`, …) followed
/// by the per-lane values, laid out little-endian into the 16-byte vector — the
/// exact layout the VM's `V128_CONST` opcode reads back.
fn encode_v128_const(args: &[&Expression]) -> [u8; 16] {
    let mut bytes = [0u8; 16];
    let shape = match args.first().map(|e| &e.kind) {
        Some(ExprKind::Lit(Literal::Str(s))) => s.as_str(),
        _ => "i32x4",
    };
    let vals = &args[1..];
    let int_at = |i: usize| -> i64 {
        match vals.get(i).map(|e| &e.kind) {
            Some(ExprKind::Lit(Literal::Int(n))) => *n,
            _ => 0,
        }
    };
    let float_at = |i: usize| -> f64 {
        match vals.get(i).map(|e| &e.kind) {
            Some(ExprKind::Lit(Literal::Float(f))) => *f,
            Some(ExprKind::Lit(Literal::Int(n))) => *n as f64,
            _ => 0.0,
        }
    };
    match shape {
        "i8x16" => {
            for (i, b) in bytes.iter_mut().enumerate() {
                *b = int_at(i) as u8;
            }
        }
        "i16x8" => {
            for i in 0..8 {
                bytes[i * 2..i * 2 + 2].copy_from_slice(&(int_at(i) as u16).to_le_bytes());
            }
        }
        "i32x4" => {
            for i in 0..4 {
                bytes[i * 4..i * 4 + 4].copy_from_slice(&(int_at(i) as u32).to_le_bytes());
            }
        }
        "i64x2" => {
            for i in 0..2 {
                bytes[i * 8..i * 8 + 8].copy_from_slice(&(int_at(i) as u64).to_le_bytes());
            }
        }
        "f32x4" => {
            for i in 0..4 {
                bytes[i * 4..i * 4 + 4].copy_from_slice(&(float_at(i) as f32).to_le_bytes());
            }
        }
        "f64x2" => {
            for i in 0..2 {
                bytes[i * 8..i * 8 + 8].copy_from_slice(&float_at(i).to_le_bytes());
            }
        }
        _ => {}
    }
    bytes
}

impl Compiler {
    // Builtins (profile-driven)
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn try_compile_builtin(
        &mut self,
        name: &str,
        args: &[&Expression],
    ) -> Result<bool, String> {
        let line = self.line;

        if self.is_python_profile() && name == "globals" && args.is_empty() {
            common::dict::emit_new(&mut self.chunks, self.current, line);

            inst!(self, core_wasm::dup);
            self.emit_const(Value::String(Arc::from("__main__")));
            let name_key = self.str_const("__name__");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);
            inst!(self, core_wasm::dup);
            let keys_key = self.str_const("__keys");
            self.emit_u16(Op::STRUCT_GET, keys_key);
            self.emit_const(Value::String(Arc::from("__name__")));
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);

            let mut globals: Vec<String> = self.defined_globals.iter().cloned().collect();
            globals.sort();
            globals.dedup();
            for global in globals {
                if global == "__name__" {
                    continue;
                }
                inst!(self, core_wasm::dup);
                self.emit_var_get(&global);
                let key = self.str_const(&global);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);

                inst!(self, core_wasm::dup);
                let keys_key = self.str_const("__keys");
                self.emit_u16(Op::STRUCT_GET, keys_key);
                self.emit_const(Value::String(Arc::from(global.as_str())));
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
            }
            return Ok(true);
        }

        if self.is_python_profile() && name == "frozenset" && args.len() <= 1 {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                let idx = self.import("ecma:array", "from");
                self.emit_host_call(idx, 1);
                self.emit_const(Value::String(Arc::from("\u{1f}")));
                common::collections::emit_join(&mut self.chunks, self.current, line);
            } else {
                self.emit_const(Value::String(Arc::from("")));
            }
            return Ok(true);
        }

        if self.is_php_profile() {
            let builtin_name = self.canon(name);
            if builtin_name == "strval" && args.len() == 1 {
                self.compile_expr(args[0])?;
                self.emit_common("php.echo_stringify", 1, line);
                return Ok(true);
            }

            if builtin_name == "intval" && args.len() == 2 {
                self.compile_expr(args[0])?;
                self.compile_expr(args[1])?;
                let parse_int = self.import("ecma:number", "parseInt");
                self.emit_host_call(parse_int, 2);
                return Ok(true);
            }
        }

        if self.profile.name == "pascal" {
            let builtin_name = self.canon(name);
            if builtin_name == "write" || builtin_name == "writeln" {
                let mut part_count = 0usize;
                for (index, arg) in args.iter().enumerate() {
                    if index > 0 {
                        self.emit_const(Value::String(Arc::from(" ")));
                        part_count += 1;
                    }
                    self.compile_expr(arg)?;
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);
                    part_count += 1;
                }
                if builtin_name == "writeln" {
                    self.emit_const(Value::String(Arc::from("\n")));
                    part_count += 1;
                }
                let line = self.line;
                common::strings::emit_concat(self.chunk(), part_count, line);

                let text_slot = self.define_local("__pascal_stdout_text");
                self.emit_u16(Op::LOCAL_SET, text_slot);

                let stdout_idx = self.import("wasi:cli/stdout", "get-stdout");
                let write_idx = self.import(
                    "wasi:io/streams",
                    "[method]output-stream.blocking-write-and-flush",
                );
                self.emit_host_call(stdout_idx, 0);
                self.emit_u16(Op::LOCAL_GET, text_slot);
                self.emit_host_call(write_idx, 2);
                self.emit(Op::DROP);
                return Ok(true);
            }

            if (builtin_name == "integer" || builtin_name == "int" || builtin_name == "longint")
                && args.len() == 1
            {
                self.compile_expr(args[0])?;
                common::math::emit_trunc(self.chunk(), line);
                return Ok(true);
            }

            if builtin_name == "inttohex" && (1..=2).contains(&args.len()) {
                self.compile_expr(args[0])?;
                let number_idx = self.import("ecma:number", "Number");
                self.emit_host_call(number_idx, 1);
                self.emit_const(Value::F64(16.0));
                let to_string_idx = self.import("ecma:number", "toString");
                self.emit_host_call(to_string_idx, 2);
                let upper_idx = self.import("ecma:string", "toUpperCase");
                self.emit_host_call(upper_idx, 1);
                if let Some(width) = args.get(1) {
                    self.compile_expr(width)?;
                    self.emit_const(Value::String(Arc::from("0")));
                    let pad_start_idx = self.import("ecma:string", "padStart");
                    self.emit_host_call(pad_start_idx, 3);
                }
                return Ok(true);
            }

            if builtin_name == "booltostr" && (1..=2).contains(&args.len()) {
                self.compile_expr(args[0])?;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                self.emit_const(Value::String(Arc::from(if args.len() == 1 {
                    "true"
                } else {
                    "True"
                })));
                self.chunk().emit_else(line);
                self.emit_const(Value::String(Arc::from(if args.len() == 1 {
                    "false"
                } else {
                    "False"
                })));
                self.chunk().emit_end(line);
                return Ok(true);
            }

            if (builtin_name == "ansiuppercase" || builtin_name == "ansilowercase")
                && args.len() == 1
            {
                self.compile_expr(args[0])?;
                let method = if builtin_name == "ansiuppercase" {
                    "toUpperCase"
                } else {
                    "toLowerCase"
                };
                let idx = self.import("ecma:string", method);
                self.emit_host_call(idx, 1);
                return Ok(true);
            }

            if builtin_name == "samestr" && args.len() == 2 {
                self.compile_expr(args[0])?;
                self.compile_expr(args[1])?;
                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                return Ok(true);
            }

            if (builtin_name == "sametext" || builtin_name == "comparetext") && args.len() == 2 {
                self.compile_expr(args[0])?;
                let lower_idx = self.import("ecma:string", "toLowerCase");
                self.emit_host_call(lower_idx, 1);
                self.compile_expr(args[1])?;
                self.emit_host_call(lower_idx, 1);
                if builtin_name == "sametext" {
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                } else {
                    let compare_idx = self.import("ecma:string", "localeCompare");
                    self.emit_host_call(compare_idx, 2);
                }
                return Ok(true);
            }

            if builtin_name == "strtobool" && args.len() == 1 {
                self.compile_expr(args[0])?;
                let lower_idx = self.import("ecma:string", "toLowerCase");
                self.emit_host_call(lower_idx, 1);
                self.emit_const(Value::String(Arc::from("true")));
                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                return Ok(true);
            }

            if builtin_name == "strtointdef" && args.len() == 2 {
                self.compile_expr(args[0])?;
                let parse_idx = self.import("ecma:number", "parseInt");
                self.emit_host_call(parse_idx, 1);
                let parsed_slot = self.define_local("__pascal_strtointdef_value");
                self.emit_u16(Op::LOCAL_SET, parsed_slot);
                self.emit_u16(Op::LOCAL_GET, parsed_slot);
                let is_nan_idx = self.import("ecma:number", "isNaN");
                self.emit_host_call(is_nan_idx, 1);
                self.chunk().emit_if_value(line);
                self.compile_expr(args[1])?;
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, parsed_slot);
                self.chunk().emit_end(line);
                return Ok(true);
            }

            if builtin_name == "delete"
                && args.len() == 3
                && matches!(&args[0].kind, ExprKind::Ident(_))
            {
                let ExprKind::Ident(var_name) = &args[0].kind else {
                    unreachable!();
                };
                let helper_idx = self.str_const("__vybe_pascal_str_remove_range");
                self.emit_u16(Op::GLOBAL_GET, helper_idx);
                self.emit_var_get(var_name);
                self.compile_expr(args[1])?;
                self.compile_expr(args[2])?;
                self.emit_u8(Op::CALL_REF, 3);
                self.emit_var_set(var_name);
                self.emit(Op::NULL);
                return Ok(true);
            }

            if builtin_name == "insert"
                && args.len() == 3
                && matches!(&args[1].kind, ExprKind::Ident(_))
            {
                let ExprKind::Ident(var_name) = &args[1].kind else {
                    unreachable!();
                };
                let helper_idx = self.str_const("__vybe_pascal_str_insert");
                self.emit_u16(Op::GLOBAL_GET, helper_idx);
                self.compile_expr(args[0])?;
                self.emit_var_get(var_name);
                self.compile_expr(args[2])?;
                self.emit_u8(Op::CALL_REF, 3);
                self.emit_var_set(var_name);
                self.emit(Op::NULL);
                return Ok(true);
            }
        }

        if self.profile.name == "pascal"
            && args.len() == 2
            && matches!(&args[0].kind, ExprKind::Ident(_))
        {
            let builtin_name = self.canon(name);
            let ExprKind::Ident(var_name) = &args[0].kind else {
                unreachable!();
            };

            let is_set_var = self
                .lookup_var_type_hint(var_name)
                .is_some_and(Self::is_pascal_set_type_hint);
            if is_set_var && (builtin_name == "include" || builtin_name == "exclude") {
                let helper = if builtin_name == "include" {
                    "__vybe_pascal_set_include"
                } else {
                    "__vybe_pascal_set_exclude"
                };
                let helper_idx = self.str_const(helper);
                self.emit_u16(Op::GLOBAL_GET, helper_idx);
                self.emit_var_get(var_name);
                self.compile_expr(args[1])?;
                self.emit_u8(Op::CALL_REF, 2);
                self.emit(Op::DROP);
                self.emit(Op::NULL);
                return Ok(true);
            }
        }

        if name.eq_ignore_ascii_case("setlength") {
            if args.len() >= 2 {
                self.compile_setlength(args[0], args[1])?;
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }

        // ── Component Model host-call resolution (qualified name → host fn) ──
        //
        // A qualified identifier whose first segment matches the profile's
        // `host_packages` list resolves directly to a Component Model host
        // call. This is how `\Vybe\Http\Response\set_status(404)` in PHP
        // reaches the `vybe:http/response` host module with zero profile
        // builtin entries. The same convention is intended to apply to every
        // language with namespaces (Python `vybe.http.request.method`, C#
        // `Vybe.Http.Request.Method`, etc.) — walkers normalize their
        // separators to `\` before reaching here so this single resolver
        // handles them all.
        if let Some((module, func)) = self.resolve_component_model_call(name) {
            for a in args {
                self.compile_expr(a)?;
            }
            let idx = self.import(&module, &func);
            self.emit_host_call(idx, args.len() as u8);
            return Ok(true);
        }

        // ── Phase D1 pilot: Array(count, init) → ecma:array.newWithLength + fill ──
        //
        // COBOL's OCCURS walker emits `Call { callee: Array,
        // args: [count, element_init] }` in the high-level IR. This
        // intercept routes the pattern through the spec-conformant
        // `ecma:array.*` imports instead of the legacy VM-internal
        // opcodes. See `dynamicruntime_support.md` Phase D1 and the
        // reasoning in `project_dynamic_runtime_phase_state.md`.
        //
        // Narrow match: only intercept when we see `Array(count, init)`
        // specifically — 2 positional args, callee identifier "Array".
        // This avoids colliding with C#/VB `Array` namespace access
        // (`Array.Empty()`, `Array.IsArray()`, etc.) which hits
        // different code paths (namespace + member access).
        if name == "Array" && args.len() == 2 {
            // COBOL's OCCURS walker emits `Array(count, init)`. Emit:
            //   newWithLength(count)  — via common::collections
            //   fill(arr, init, 0, MAX)  — via common::collections
            self.compile_expr(args[0])?; // push count
            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            // Array is now on TOS. If the init is null-ish, we're done
            // (newWithLength already null-fills).
            let init_is_null = matches!(
                &args[1].kind,
                ExprKind::Lit(crate::ast::Literal::Null)
                    | ExprKind::Lit(crate::ast::Literal::Undefined)
            );
            if init_is_null {
                return Ok(true);
            }
            let init_is_nested_array_factory = matches!(
                &args[1].kind,
                ExprKind::Call { callee, .. }
                    if matches!(callee.kind, ExprKind::Ident(ref name) if name == "Array")
            );
            if init_is_nested_array_factory {
                let arr_slot = self.define_local("__array_ctor_result");
                self.emit_u16(Op::LOCAL_SET, arr_slot);

                let idx_slot = self.define_local("__array_ctor_idx");
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, idx_slot);

                let fill_block = self.chunk().emit_block(line);
                let (fill_loop, _) = self.chunk().emit_loop_s(line);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, arr_slot);
                common::collections::emit_len(&mut self.chunks, self.current, line);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
                self.chunk().emit_br_if(1, line);

                self.emit_u16(Op::LOCAL_GET, arr_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.compile_expr(args[1])?;
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_const(Value::I32(1));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                };
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.chunk().emit_br(0, line);
                self.chunk().emit_end(line);
                self.chunk().patch_loop(fill_loop);
                self.chunk().emit_end(line);
                self.chunk().patch_block(fill_block);
                self.emit_u16(Op::LOCAL_GET, arr_slot);
                return Ok(true);
            }
            // Stack: [arr]. Dup first so we still have the result.
            inst!(self, core_wasm::dup);
            self.compile_expr(args[1])?;
            inst!(self, core_wasm::i32_const, 0);
            inst!(self, core_wasm::i32_const, i32::MAX);
            common::collections::emit_fill(&mut self.chunks, self.current, line);
            // fill returns the array; drop the dup'd copy — the pre-dup
            // copy stays on TOS as the expression's value.
            self.emit(Op::DROP);
            return Ok(true);
        }

        if name.eq_ignore_ascii_case("__fortran_max") {
            if args.len() >= 2 {
                self.compile_expr(args[0])?;
                self.compile_expr(args[1])?;
                common::math::emit_max(self.chunk(), line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("__fortran_min") {
            if args.len() >= 2 {
                self.compile_expr(args[0])?;
                self.compile_expr(args[1])?;
                common::math::emit_min(self.chunk(), line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("__fortran_emit")
            || name.eq_ignore_ascii_case("__fortran_emitln")
        {
            let flush = name.eq_ignore_ascii_case("__fortran_emitln");
            let text_slot = self.define_local(if flush {
                "__fortran_emitln_text"
            } else {
                "__fortran_emit_text"
            });
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
            } else {
                self.compile_expr(&Expression::string(""))?;
            }
            self.emit_u16(Op::LOCAL_SET, text_slot);

            self.emit_var_get("__vybe_fortran_io_buffer");
            self.emit_u16(Op::LOCAL_GET, text_slot);
            common::strings::emit_str_concat(self.chunk(), line);

            if flush {
                let message_slot = self.define_local("__fortran_emitln_message");
                self.emit_u16(Op::LOCAL_SET, message_slot);

                self.emit_u16(Op::LOCAL_GET, message_slot);
                let idx = self.import("wasi:logging/logging", "log");
                common::io::emit_print_with_import(self.chunk(), idx, 1, line);

                self.compile_expr(&Expression::string(""))?;
                self.emit_var_set("__vybe_fortran_io_buffer");
            } else {
                self.emit_var_set("__vybe_fortran_io_buffer");
                self.emit(Op::NULL);
            }
            return Ok(true);
        }

        if name.eq_ignore_ascii_case("__fortran_rewind") {
            let file_slot = self.define_local("__fortran_rewind_file");
            let path_slot = self.define_local("__fortran_rewind_path");

            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
            } else {
                self.emit_const(Value::I32(0));
            }
            self.emit_u16(Op::LOCAL_SET, file_slot);

            self.emit_global_map_get_into_local("__vb_file_path_by_handle", file_slot, path_slot);

            self.emit_u16(Op::LOCAL_GET, file_slot);
            let close_idx = self.import("wasi:filesystem", "closeFile");
            self.emit_host_call(close_idx, 1);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, path_slot);
            self.emit_const(Value::String(Arc::from("Input")));
            self.emit_u16(Op::LOCAL_GET, file_slot);
            let open_idx = self.import("wasi:filesystem", "openFile");
            self.emit_host_call(open_idx, 3);
            self.emit(Op::DROP);

            self.emit_global_map_set_const(
                "__vb_file_eof_by_handle",
                file_slot,
                Value::Bool(false),
            );
            self.emit_global_map_set_null("__vb_record_rows_by_handle", file_slot);
            self.emit_global_map_set_null("__vb_record_next_index_by_handle", file_slot);
            self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
            self.emit(Op::NULL);
            return Ok(true);
        }

        if name.eq_ignore_ascii_case("__fortran_namelist_decl") {
            self.emit(Op::NULL);
            return Ok(true);
        }

        if name.eq_ignore_ascii_case("kind") {
            self.emit_const(Value::I32(8));
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("allocate") {
            for arg in args {
                match &arg.kind {
                    ExprKind::Call {
                        callee, args: dims, ..
                    } if !dims.is_empty() => {
                        if self.profile.name == "fortran" {
                            let mut dim_slots = Vec::with_capacity(dims.len());
                            for (index, dim) in dims.iter().enumerate() {
                                self.compile_expr(&dim.value)?;
                                let slot =
                                    self.define_local(&format!("__fortran_alloc_dim_{index}"));
                                self.emit_u16(Op::LOCAL_SET, slot);
                                dim_slots.push(slot);
                            }

                            let ctor_name = self.fortran_allocate_ctor_name(callee);
                            self.emit_fortran_allocated_array(&dim_slots, ctor_name.as_deref());
                            self.compile_assign_target(callee)?;
                        } else {
                            self.compile_expr(&dims[0].value)?;
                            common::collections::emit_new_with_length(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            self.compile_assign_target(callee)?;
                        }
                    }
                    _ => {
                        if self.profile.name == "fortran" {
                            if let Some(ctor_name) = self.fortran_allocate_ctor_name(arg) {
                                self.emit_fortran_ctor_call(&ctor_name);
                            } else {
                                self.emit(Op::NULL);
                            }
                            self.compile_assign_target(arg)?;
                        }
                    }
                }
            }
            self.emit(Op::NULL);
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("deallocate") {
            for arg in args {
                match &arg.kind {
                    ExprKind::Call { callee, .. } => {
                        self.emit(Op::NULL);
                        self.compile_assign_target(callee)?;
                    }
                    _ => {
                        self.emit(Op::NULL);
                        self.compile_assign_target(arg)?;
                    }
                }
            }
            self.emit(Op::NULL);
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("present") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                self.emit(Op::REF_IS_NULL);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            } else {
                inst!(self, core_wasm::bool_const, false);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("sum") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::collections::emit_sum(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("minval") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::collections::emit_pymin(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("maxval") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::collections::emit_pymax(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("nint") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::math::emit_round(self.chunk(), line);
                common::convert::emit_to_int(self.chunk(), line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if !self.profile.has_ecma_globals && name.eq_ignore_ascii_case("size") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::collections::emit_len(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("matmul") {
            for arg in args {
                self.compile_expr(arg)?;
            }
            self.emit_common("fortran.matmul", args.len() as u8, line);
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("array_join") {
            if args.len() >= 2 {
                self.compile_expr(args[0])?;
                self.compile_expr(args[1])?;
                common::collections::emit_join(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("str_getcsv") {
            for arg in args {
                self.compile_expr(arg)?;
            }
            vybe_runtime::registry::hooks(&self.profile.name)
                .str_getcsv
                .unwrap()(&mut self.chunks, self.current, args.len() as u8, line);
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("this_image") {
            self.emit_const(Value::I32(1));
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("num_images") {
            self.emit_const(Value::I32(1));
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("co_sum") {
            self.emit(Op::NULL);
            return Ok(true);
        }

        // Canonical builtins — language-agnostic dispatch via compiler_common::canonical.
        // Walkers normalize language-specific syntax (arr.Length, len(arr), Length(arr),
        // arr.size, etc.) to canonical dunder names (__len__, __str__, etc.).
        // The compiler doesn't know about language-specific names — it just looks up
        // the canonical name in compiler_common's registry.
        if let Some(canonical_op) = crate::primitives::canonical::CanonicalOp::from_name(name) {
            if self.profile.has_ecma_globals
                && matches!(canonical_op, crate::primitives::canonical::CanonicalOp::Len)
            {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let obj_slot = self.define_local("__js_len_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    if self.uses_proxy {
                        // `.length` must fire the proxy get trap like any
                        // other member read (the dotted form normalizes to
                        // __len__ and would otherwise bypass §10.5.8).
                        self.emit_const(Value::String(Arc::from("length")));
                        vybe_runtime::registry::hooks(&self.profile.name)
                            .proxy_get
                            .unwrap()(&mut self.chunks, self.current, line);
                    } else {
                        let length_key = self.str_const("length");
                        self.emit_u16(Op::STRUCT_GET, length_key);
                    }
                    // §10.1.8.1 OrdinaryGet: a missing own `length` walks
                    // the prototype chain like any other key (e.g.
                    // AsyncFunction.prototype.length inherits
                    // %Function.prototype%'s 0).
                    let val_slot = self.define_local("__js_len_val");
                    self.emit_u16(Op::LOCAL_SET, val_slot);
                    self.emit_u16(Op::LOCAL_GET, val_slot);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    self.chunk().emit_if_value(line);
                    self.emit_member_get_from_value(obj_slot, "length");
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, val_slot);
                    self.chunk().emit_end(line);
                    return Ok(true);
                }
            }
            // Special case: __str__ uses stdlib via global, not host import
            if matches!(canonical_op, crate::primitives::canonical::CanonicalOp::Str) {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let arg_slot = self.define_local("__canonical_str_arg");
                    self.emit_u16(Op::LOCAL_SET, arg_slot);

                    self.emit_u16(Op::LOCAL_GET, arg_slot);
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);
                    return Ok(true);
                }
            } else {
                // Compile args, then dispatch to canonical emitter
                for a in args {
                    self.compile_expr(a)?;
                }
                crate::primitives::canonical::emit_canonical(
                    canonical_op,
                    &mut self.chunks,
                    self.current,
                    line,
                );
                return Ok(true);
            }
        }

        // Look up in language profile FIRST — language profiles can
        // override the common import defaults (e.g. Dart `print` needs
        // toString conversion before logging, which is different from
        // generic `wasi:cli.log`).
        // DateTime field-extractor functions (VB `Year(d)` → `d.Year`) are
        // declared per-language in `[datetime_field_functions]`; the DateTime
        // type check stays here (compile-time analysis). No language-name gate.
        if args.len() == 1 {
            if let Some(field_name) = self
                .profile
                .datetime_field_functions
                .get(&name.to_lowercase())
                .cloned()
            {
                if let Some(type_hint) = self.infer_expr_type_hint(&args[0]) {
                    let normalized = Self::normalize_type_hint(&type_hint);
                    if normalized == "datetime" || normalized.ends_with(".datetime") {
                        self.compile_expr(&args[0])?;
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_GET, idx);
                        return Ok(true);
                    }
                }
            }
        }

        let builtin = self.profile.lookup_builtin(name).cloned();
        // Check common import table only if the profile didn't bind it.
        if builtin.is_none() {
            if let Some(resolved) = crate::primitives::imports::resolve_common_import(name) {
                match resolved {
                    crate::primitives::imports::CommonImport::Host(module, func) => {
                        for a in args {
                            self.compile_expr(a)?;
                        }
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, args.len() as u8);
                    }
                    crate::primitives::imports::CommonImport::Intrinsic(intrinsic_name) => {
                        self.emit_intrinsic(intrinsic_name, args)?;
                    }
                }
                return Ok(true);
            }
        }

        // builtinslotplan.md step 3 — the SECOND census site. The first one
        // (`calls.rs`, value-method dispatch) only ever sees `obj.method()`,
        // so it recorded zero rows for PHP and none of Python's `len()`: those
        // languages spell their built-in operations as FREE FUNCTIONS
        // (`strlen($s)`, `count($xs)`, `len(s)`) and arrive here instead.
        // Deriving step 5's flip list from the method-shaped site alone would
        // have read "PHP has no built-in slot traffic", which is an artefact of
        // where the hook sat, not a fact about PHP.
        //
        // Argument 0 is the receiver for this shape. Emits nothing; off unless
        // VYBE_SLOT_AUDIT is set.
        // builtinslotplan.md step 5, free-function shape. Languages that spell a
        // built-in operation as a call rather than a method (`strlen($s)`,
        // `len(s)`, Dart's desugared `__dart_index_get(recv, i)`) reach dispatch
        // here, so the binding has to apply here too or those languages could
        // never be flipped.
        //
        // Argument 0 is the receiver for this shape — the same assumption the
        // census makes, and the same limitation: it is wrong for PHP builtins
        // whose operand is not first (`str_replace(…, $subject)`). A profile
        // only gets the substitution if it declared `slot`, so that limitation
        // is opt-in per method rather than applied blindly.
        let builtin = match (builtin, args.first()) {
            (Some(def), Some(receiver)) => Some(self.apply_builtin_slot_binding(receiver, def)),
            (other, _) => other,
        };

        // Census AFTER the substitution, matching the value-method site. It
        // reports the EFFECTIVE emit — the one actually compiled — so the two
        // sites answer the same question. Logging the pre-substitution emit
        // here made a binding that had fired look like it had not.
        if let (Some(def), Some(receiver)) = (builtin.as_ref(), args.first()) {
            self.audit_builtin_slot_census(receiver, name, &def.emit);
        }

        if let Some(def) = builtin {
            match &def.emit {
                BuiltinEmit::Print => {
                    if self.is_php_profile() && name.eq_ignore_ascii_case("var_dump") {
                        let idx = self.import("wasi:logging/logging", "log");
                        for a in args {
                            self.compile_expr(a)?;
                            self.emit_common("php.var_dump_stringify", 1, line);
                            common::io::emit_print_with_import(self.chunk(), idx, 1, line);
                        }
                        return Ok(true);
                    }
                    // print / console.log → `wasi:logging/logging.log`. The
                    // host log fn renders each arg via the console/inspect
                    // surface (`Value::Display`: BigInt `8n`, `-0`, arrays
                    // `1,2`, …) — NOT ECMAScript `ToString`. Keep the
                    // stringification in the host so the ECMA console base
                    // stays spec-correct; dotnet layers its own formatting on
                    // top via `emit_dotnet_console_arg`.
                    let mut arg_slots = Vec::with_capacity(args.len());
                    for (index, a) in args.iter().enumerate() {
                        if let Some((_, member_name)) = self.qualified_enum_member_expr(a) {
                            self.emit_const(Value::String(Arc::from(member_name.as_str())));
                        } else if let Some(enum_type) = self.console_enum_type_from_expr(a) {
                            self.emit_enum_value_to_string(&enum_type, a)?;
                        } else {
                            self.compile_expr(a)?;
                        }
                        let arg_slot = self.define_local(&format!("__print_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                    for slot in &arg_slots {
                        self.emit_u16(Op::LOCAL_GET, *slot);
                    }
                    let idx = self.import("wasi:logging/logging", "log");
                    common::io::emit_print_with_import(self.chunk(), idx, args.len() as u8, line);
                }
                BuiltinEmit::StrLength => {
                    if !args.is_empty() {
                        self.compile_expr(args[0])?;
                        common::strings::emit_length(self.chunk(), line);
                    } else {
                        self.emit(Op::NULL);
                    }
                }
                BuiltinEmit::HostCall(module, func) => {
                    // Iterator-consuming host fns (e.g. `Array.from`,
                    // `Promise.all`) accept any iterable. JS generators
                    // (Continuation) need WASM stack-switching to
                    // drain — a host fn can't drive coroutine resume,
                    // so we drain via the `__stdlib_drain_generator`
                    // bytecode helper before the host call.
                    let drains_sync_iterable = (module == "ecma:array" && func == "from")
                        || (module == "ecma:iterator" && func == "from")
                        || (module == "ecma:object" && func == "fromEntries")
                        || (module == "ecma:object" && func == "groupBy")
                        || (module == "ecma:map" && func == "groupBy")
                        || (matches!(
                            module.as_str(),
                            "ecma:int8array"
                                | "ecma:uint8array"
                                | "ecma:uint8clamped"
                                | "ecma:int16array"
                                | "ecma:uint16array"
                                | "ecma:int32array"
                                | "ecma:uint32array"
                                | "ecma:float32array"
                                | "ecma:float64array"
                                | "ecma:bigint64array"
                                | "ecma:biguint64array"
                        ) && func == "from")
                        || (module == "ecma:promise"
                            && matches!(func.as_str(), "all" | "allSettled" | "race" | "any"));
                    let drains_async_iterable = (module == "ecma:array" && func == "fromAsync")
                        || (module == "ecma:iterator" && func == "asyncFrom");
                    let drain_first_arg = self.profile.has_generators
                        && (drains_sync_iterable || drains_async_iterable);
                    let async_drain = self.profile.has_generators && drains_async_iterable;
                    if drain_first_arg && !args.is_empty() {
                        let rejects_primitive_iterable_arg = (module == "ecma:promise"
                            && matches!(func.as_str(), "all" | "allSettled" | "race" | "any"))
                            || ((module == "ecma:object" || module == "ecma:map")
                                && func == "groupBy");
                        let rejects_nullish_source = matches!(
                            module.as_str(),
                            "ecma:int8array"
                                | "ecma:uint8array"
                                | "ecma:uint8clamped"
                                | "ecma:int16array"
                                | "ecma:uint16array"
                                | "ecma:int32array"
                                | "ecma:uint32array"
                                | "ecma:float32array"
                                | "ecma:float64array"
                                | "ecma:bigint64array"
                                | "ecma:biguint64array"
                        ) && func == "from";
                        if rejects_primitive_iterable_arg
                            && matches!(
                                args[0].kind,
                                ExprKind::Lit(
                                    Literal::Int(_)
                                        | Literal::Float(_)
                                        | Literal::BigInt(_)
                                        | Literal::Bool(_)
                                        | Literal::Char(_)
                                        | Literal::Null
                                        | Literal::Undefined
                                        | Literal::Ellipsis
                                )
                            )
                        {
                            let label = if module == "ecma:promise" {
                                format!("Promise.{func}")
                            } else if module == "ecma:object" {
                                "Object.groupBy".to_string()
                            } else {
                                "Map.groupBy".to_string()
                            };
                            self.emit_const(Value::String(Arc::from(format!(
                                "{label} argument is not iterable"
                            ))));
                            self.emit_js_exception_ctor_from_message_value("TypeError")?;
                            let line = self.line;
                            common::errors::emit_throw(self.chunk(), line);
                            return Ok(true);
                        }
                        if rejects_nullish_source
                            && matches!(
                                args[0].kind,
                                ExprKind::Lit(Literal::Null | Literal::Undefined)
                            )
                        {
                            self.emit_const(Value::String(Arc::from(
                                "TypedArray.from source is not iterable",
                            )));
                            self.emit_js_exception_ctor_from_message_value("TypeError")?;
                            let line = self.line;
                            common::errors::emit_throw(self.chunk(), line);
                            return Ok(true);
                        }
                        self.compile_expr(args[0])?;
                        if async_drain {
                            common::generators::emit_drain_async_iterable(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                        } else {
                            common::collections::emit_spread_iterable(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                        }
                        for a in args.iter().skip(1) {
                            self.compile_expr(a)?;
                        }
                    } else {
                        for a in args {
                            self.compile_expr(a)?;
                        }
                    }
                    let idx = self.import(module, func);
                    self.emit_host_call(idx, args.len() as u8);
                }
                BuiltinEmit::Opcode(op_name) => {
                    self.emit_builtin_opcode(op_name, args)?;
                }
                BuiltinEmit::MutateVar(op) => {
                    if let Some(first) = args.first() {
                        if let ExprKind::Ident(var) = &first.kind {
                            let var = var.clone();
                            self.emit_var_get(&var);
                            if args.len() > 1 {
                                self.compile_expr(args[1])?;
                            } else {
                                self.emit_const(Value::F64(1.0));
                            }
                            match op.as_str() {
                                "add" => {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                                }
                                "sub" => self.emit(Op::F64_SUB),
                                _ => {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                                }
                            }
                            self.emit_var_set(&var);
                        }
                    }
                    self.emit(Op::NULL);
                }
                BuiltinEmit::Intrinsic(intrinsic_name) => {
                    self.emit_intrinsic(intrinsic_name, args)?;
                }
                BuiltinEmit::Common(name) => {
                    // Compile args, then dispatch to compiler_common emitter.
                    // Console.WriteLine/Write should preserve enum names instead
                    // of writing raw ordinals, and apply .NET numeric formatting.
                    if (name.eq_ignore_ascii_case("dotnet.console_writeline")
                        || name.eq_ignore_ascii_case("dotnet.console_write"))
                        && args.len() == 1
                    {
                        self.emit_dotnet_console_arg(args[0])?;
                    } else {
                        for a in args {
                            self.compile_expr(a)?;
                        }
                    }
                    let line = self.line;
                    self.emit_common(name.as_str(), args.len() as u8, line);
                }
                BuiltinEmit::Noop => {
                    self.emit(Op::NULL);
                }
                BuiltinEmit::Invoke(_) => {
                    // `invoke:` is only meaningful for value-method calls
                    // (receiver in hand). In the free-function path the
                    // profile shouldn't use it — emit null so misconfigured
                    // profiles fail loudly via type checks rather than
                    // silent wrong behaviour.
                    self.emit(Op::NULL);
                }
            }
            return Ok(true);
        }

        // Generic 1:1 opcode route (WAT/WAST only): the wast profile lowers every
        // instruction it doesn't special-case to `Call(instr_name_underscored)`,
        // so a call whose name is a WASM instruction with no higher-level builtin
        // (e.g. every SIMD lane op) emits that opcode directly. The single opcode
        // list lives in the VM — `Op::from_wasm_name` resolves it; there is no
        // per-op list in the profile or the emitter. Gated on the
        // `function_references` property (set only by raw-WASM frontends, i.e.
        // wast) so bare opcode-shaped identifiers (`select`, `drop`, …) in other
        // languages are never mistaken for instructions.
        if self.profile.function_references {
            // WASM GC array ops. `array.new`/`array.new_default` stamp the
            // instance with the registry id of a real `(array …)` defined type
            // so the VM applies spec trapping `array.get`/`set`/`copy` (null /
            // out-of-bounds) — the same rtt mechanism `ref.test`/`ref.cast` use.
            // The null check for get/set stays compiler-side (a typed null
            // carries no instance to stamp); the VM enforces the bounds trap.
            match name {
                "array.new" | "array_new" => {
                    self.emit_gc_array_new(args, false)?;
                    return Ok(true);
                }
                "array.new_default" | "array_new_default" => {
                    self.emit_gc_array_new(args, true)?;
                    return Ok(true);
                }
                // Compile-time directive from the wast walker: install a GC
                // struct type (with its parent = subtype edge) in the type
                // table so `ref.test`/`ref.cast`/`br_on_cast` resolve identity
                // and subtyping. Emits no runtime code.
                "__wast_register_struct_type" => {
                    self.register_wast_struct_type(args);
                    return Ok(true);
                }
                // Compile-time directive from the wast walker: install a GC
                // `(array …)` type carrying its element storage type so the VM
                // can recover the element byte width from an instance's rtt.
                // Emitted before any `array.*` so `resolve_gc_array_type_id`
                // finds this (element-typed) entry rather than a bare one.
                "__wast_register_array_type" => {
                    let name = expr_str_lit(args.first().copied());
                    let elem = expr_str_lit(args.get(1).copied());
                    if !name.is_empty() {
                        let key = format!("__wast_array::{name}");
                        if let Some(idx) = self.chunks[0].types.iter().position(|t| t.name == key) {
                            if self.chunks[0].types[idx].fields.is_empty() && !elem.is_empty() {
                                self.chunks[0].types[idx].fields = vec![elem];
                            }
                        } else {
                            crate::primitives::classes::register_gc_array_type(
                                &mut self.chunks,
                                &key,
                                &elem,
                            );
                        }
                    }
                    return Ok(true);
                }
                // Compile-time directive from the wast walker: register a passive
                // element segment's funcref list (resolved to function chunk
                // indices) under its segment index, so `table.init`/`array.new_elem`
                // read real funcrefs the VM materializes at instantiation.
                "__wast_register_passive_elem" => {
                    let seg_index = expr_const_u16(args.first().copied()) as usize;
                    let mut chunk_indices = Vec::new();
                    for a in &args[1..] {
                        let name = expr_str_lit(Some(a));
                        if let Some(idx) = self
                            .resolve_unique_static_method_chunk_for_class("__wasm_module", &name)
                        {
                            chunk_indices.push(idx);
                        }
                    }
                    if self.chunks[0].passive_elem_funcs.len() <= seg_index {
                        self.chunks[0]
                            .passive_elem_funcs
                            .resize_with(seg_index + 1, Vec::new);
                    }
                    self.chunks[0].passive_elem_funcs[seg_index] = chunk_indices;
                    return Ok(true);
                }
                // `ref.test <ht>` / `ref.cast <ht>` — runtime type check against
                // the named GC type via the registered hierarchy. A nullable
                // heap type (`(ref null $T)`) uses the `_NULL` op (null passes);
                // a bare `$T` uses the non-null op (null → test 0 / cast trap).
                // Null guard for WASM GC `struct.get`/`struct.set`: pass the ref
                // through, or trap if it is null (spec). Wrapped around the field
                // access object by the walker so the read/write lowering is
                // unchanged; dynamic-language member access is never routed here.
                // `ref.null $t` → a WASM GC typed null: the single `ref.null none`
                // op (Op::NULL_NONE) pushes a TypedNull and round-trips as
                // `0xD0 0x71` in `.wasm`. The VM's GC accessors trap on it; it's a
                // plain null elsewhere. Non-wast callers keep emitting bare
                // Op::NULL → plain null (`ref.null extern`).
                "__wast_typed_null" => {
                    let l = self.line;
                    self.chunk().emit_op(Op::NULL_NONE, l);
                    return Ok(true);
                }
                // Stamp a struct instance with its WASM GC rtt: compile the
                // object, then `GLOBAL_GET __tid_<T>` (the registered type id,
                // installed at load) + `SET_TYPE_ID` so the instance carries the
                // real `type_id` for `ref.test`/`ref.cast` — no `__type` string.
                "__wast_stamp_type" => {
                    if let Some(obj) = args.first() {
                        self.compile_expr(obj)?;
                    }
                    let type_name = expr_str_lit(args.get(1).copied());
                    let l = self.line;
                    let g = self
                        .chunk()
                        .add_constant(Value::String(std::sync::Arc::from(
                            format!("__tid_{type_name}").as_str(),
                        )));
                    self.chunk().emit_op_u16(Op::GLOBAL_GET, g, l);
                    self.chunk().emit_op(Op::SET_TYPE_ID, l);
                    return Ok(true);
                }
                // Stamp an `array.new_fixed` literal with its `(array …)` rtt:
                // register the type (so its `__tid_` global exists at load) and
                // SET_TYPE_ID, so the fixed array traps on OOB like `array.new`.
                "__wast_stamp_array_type" => {
                    if let Some(obj) = args.first() {
                        self.compile_expr(obj)?;
                    }
                    let type_ref = expr_str_lit(args.get(1).copied());
                    // Registering returns the 1-based table index; the load-time
                    // `__tid___wast_array::<ref>` global carries the registry id.
                    let _ = self.resolve_gc_array_type_id(&type_ref);
                    let l = self.line;
                    let g = self
                        .chunk()
                        .add_constant(Value::String(std::sync::Arc::from(
                            format!("__tid___wast_array::{type_ref}").as_str(),
                        )));
                    self.chunk().emit_op_u16(Op::GLOBAL_GET, g, l);
                    self.chunk().emit_op(Op::SET_TYPE_ID, l);
                    return Ok(true);
                }
                "ref.test" | "ref_test" | "ref.test_null" | "ref_test_null" => {
                    self.emit_ref_type_op(args, false)?;
                    return Ok(true);
                }
                "ref.cast" | "ref_cast" | "ref.cast_null" | "ref_cast_null" => {
                    self.emit_ref_type_op(args, true)?;
                    return Ok(true);
                }
                "array_get" => {
                    for a in args {
                        self.compile_expr(a)?;
                    }
                    let l = self.line;
                    common::collections::emit_gc_array_get(&mut self.chunks, self.current, l);
                    return Ok(true);
                }
                "array_set" => {
                    for a in args {
                        self.compile_expr(a)?;
                    }
                    let l = self.line;
                    common::collections::emit_gc_array_set(&mut self.chunks, self.current, l);
                    return Ok(true);
                }
                "array_copy" => {
                    for a in args {
                        self.compile_expr(a)?;
                    }
                    let l = self.line;
                    common::collections::emit_gc_array_copy(&mut self.chunks, self.current, l);
                    return Ok(true);
                }
                _ => {}
            }
            // A `@@mem<N>` suffix (wast multi-memory selector) is not part of the
            // opcode name — check the base name so the op still routes here;
            // `emit_builtin_opcode` strips the suffix and emits the selector.
            let base_name = name.split_once("@@mem").map(|(b, _)| b).unwrap_or(name);
            if Op::from_flattened_name(base_name).is_some() {
                self.emit_builtin_opcode(name, args)?;
                return Ok(true);
            }
        }

        Ok(false)
    }

    /// Compile a WASM GC `array.new $t value length` (or `array.new_default $t
    /// length` when `default_init`). The `$t` type reference is resolved to the
    /// registry id of a registered `(array …)` defined type; the VM stamps that
    /// id onto the instance so `array.get`/`set`/`copy` trap per spec. Stack
    /// operands follow the type immediate in fold order (value then length).
    fn emit_gc_array_new(
        &mut self,
        args: &[&Expression],
        default_init: bool,
    ) -> Result<(), String> {
        let type_name = args
            .first()
            .map(|a| wasm_type_ref_name(a))
            .unwrap_or_default();
        let type_id = self.resolve_gc_array_type_id(&type_name);
        for a in &args[1..] {
            self.compile_expr(a)?;
        }
        let l = self.line;
        let op = if default_init {
            Op::ARRAY_NEW_DEFAULT
        } else {
            Op::ARRAY_NEW
        };
        self.chunk().emit_op_u16(op, type_id as u16, l);
        Ok(())
    }

    /// Install a wast GC struct type in the type table: `args` are
    /// `[name, parent, field_count]` from the walker's compile-time directive.
    /// The parent (empty = none) becomes the subtype edge the VM's `is_subtype`
    /// walks for `ref.test`/`ref.cast`. Fields are the positional `"0".."n-1"`.
    fn register_wast_struct_type(&mut self, args: &[&Expression]) {
        let name = expr_str_lit(args.first().copied());
        if name.is_empty() {
            return;
        }
        let parent = expr_str_lit(args.get(1).copied());
        let field_count = expr_const_u16(args.get(2).copied()) as usize;
        let field_names: Vec<String> = (0..field_count).map(|i| i.to_string()).collect();
        crate::primitives::classes::register_type(
            &mut self.chunks,
            &name,
            &parent,
            field_names,
            Vec::new(),
            false,
            Vec::new(),
            None,
            std::collections::HashMap::new(),
        );
    }

    /// Emit a `ref.test`/`ref.cast` against a named heap type. `args[0]` is the
    /// heap-type reference (bare `$T` → non-null; folded `(ref null $T)` →
    /// nullable); `args[1..]` is the ref operand. The type NAME is emitted as a
    /// string constant the VM resolves through the registered hierarchy.
    fn emit_ref_type_op(&mut self, args: &[&Expression], is_cast: bool) -> Result<(), String> {
        let (type_name, nullable) = wasm_heap_type_ref(args.first().copied());
        for a in &args[1..] {
            self.compile_expr(a)?;
        }
        let l = self.line;
        let op = match (is_cast, nullable) {
            (false, false) => Op::REF_TEST,
            (false, true) => Op::REF_TEST_NULL,
            (true, false) => Op::REF_CAST,
            (true, true) => Op::REF_CAST_NULL,
        };
        let cidx = self
            .chunk()
            .add_constant(Value::String(std::sync::Arc::from(type_name.as_str())));
        self.chunk().emit_op_u16(op, cidx, l);
        Ok(())
    }

    /// The `array.new` type immediate for the `(array …)` a wast `$t` reference
    /// names: a 1-based index into `chunk 0`'s type table, keyed by the
    /// reference so every `array.new $t` shares one type (WASM type identity)
    /// and registered on first use. The VM turns this index into the runtime
    /// rtt by resolving the type *name* against the registry (the host's
    /// builtin types sit ahead of the module's, so the index is not the id).
    fn resolve_gc_array_type_id(&mut self, type_ref: &str) -> usize {
        let key = format!("__wast_array::{type_ref}");
        if let Some(idx) = self.chunks[0].types.iter().position(|t| t.name == key) {
            return idx + 1;
        }
        // No element type known here (lazy registration from a bare `array.*`);
        // the wast `__wast_register_array_type` directive registers it with the
        // element storage type ahead of use.
        crate::primitives::classes::register_gc_array_type(&mut self.chunks, &key, "")
    }

    /// Emit a compiler_common operation by namespaced name.
    /// Used by both `BuiltinEmit::Common` paths.
    ///
    /// `argc` is how many caller-supplied arguments are currently on
    /// the stack at the emit site. Multi-arity emits (e.g. .NET
    /// constructors with overloaded shapes) branch on it; most emits
    /// ignore it because their stack contract is fixed.
    pub(super) fn emit_common(&mut self, name: &str, argc: u8, line: u32) {
        // First try the import-needing dispatch (sleep, etc.). It needs a
        // closure into the compiler to resolve imports against chunk[0].
        // We use a raw pointer to break the borrow of self.
        {
            let self_ptr = self as *mut Self;
            let chunk = self.chunk();
            let handled = common::dispatch::emit_common_with_imports(
                name,
                chunk,
                argc,
                line,
                |module, fname| unsafe { (*self_ptr).import(module, fname) },
            );
            if handled {
                self.sync_scope_slots_with_chunk();
                return;
            }
        }
        // Then the pure (chunk + line) common ops.
        let line2 = line;
        let handled =
            common::dispatch::emit_common(name, &mut self.chunks, self.current, argc, line2);
        if handled {
            self.sync_scope_slots_with_chunk();
        }
        if !handled {
            eprintln!("Unknown common emit: {}", name);
        }
    }

    pub(super) fn sync_scope_slots_with_chunk(&mut self) {
        let chunk_slots = self.chunks[self.current].local_count;
        if let Some(scope) = self.scopes.last_mut() {
            if scope.next_slot < chunk_slots {
                scope.next_slot = chunk_slots;
            }
        }
    }

    /// Emit a named opcode sequence for a builtin.
    /// Emit a single opcode by name. Used for value methods where args are already on stack.
    pub(super) fn emit_named_opcode(&mut self, op_name: &str) {
        let _line = self.line;
        match op_name {
            // NOTE: bare WASM opcodes (f64.abs, ref.is_null, all SIMD lane ops, …)
            // are NOT listed here. They resolve through the single VM opcode table
            // via `Op::from_wasm_name` in the `_ =>` fallback below. Only non-opcode
            // recipes (host calls, multi-op sequences) and the two VM-internal
            // conversions (I32_FROM_F64 / F64_FROM_I32, which have no WASM mnemonic)
            // live in this match.
            "i32_from_f64" => self.emit(Op::I32_FROM_F64),
            "f64_from_i32" => self.emit(Op::F64_FROM_I32),
            "dyn_eq" => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            }
            "dyn_to_bool" => {
                let line = self.line;
                if self.is_python_profile() {
                    self.emit_condition_truthiness_from_stack();
                } else {
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                }
            }
            "dyn_not" => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_not(self.chunk(), line);
            }
            "ref_is_array" => fn_call!(self, "ecma:array", "isArray", 1),
            "ref_typeof" => fn_call!(self, "ecma:value", "typeof", 1),
            "str_length" => fn_call!(self, "wasm:js-string", "length", 1),
            "str_to_upper" => fn_call!(self, "ecma:string", "toUpperCase", 1),
            "str_to_lower" => fn_call!(self, "ecma:string", "toLowerCase", 1),
            "str_trim" => fn_call!(self, "ecma:string", "trim", 1),
            "str_trim_start" => fn_call!(self, "ecma:string", "trimStart", 1),
            "str_trim_end" => fn_call!(self, "ecma:string", "trimEnd", 1),
            "str_reverse" => {
                let l = self.line;
                crate::primitives::strings::emit_str_reverse(self.chunk(), l)
            }
            "str_from_char_code" => fn_call!(self, "wasm:js-string", "fromCharCode", 1),
            "str_char_at" => fn_call!(self, "ecma:string", "charAt", 2),
            "str_char_code_at" => fn_call!(self, "wasm:js-string", "charCodeAt", 2),
            "str_starts_with" => fn_call!(self, "ecma:string", "startsWith", 2),
            "str_ends_with" => fn_call!(self, "ecma:string", "endsWith", 2),
            "str_index_of" => fn_call!(self, "ecma:string", "indexOf", 2),
            "str_last_index_of" => fn_call!(self, "ecma:string", "lastIndexOf", 2),
            "str_includes" => {
                // includes → indexOf then check >= 0
                fn_call!(self, "ecma:string", "indexOf", 2);
                inst!(self, core_wasm::i32_const, 0);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                };
            }
            "str_contains" => fn_call!(self, "ecma:string", "includes", 2),
            "str_substring" => fn_call!(self, "wasm:js-string", "substring", 3),
            "str_split" => fn_call!(self, "ecma:string", "split", 2),
            "str_replace" => fn_call!(self, "ecma:string", "replace", 3),
            "str_repeat" => fn_call!(self, "ecma:string", "repeat", 2),
            "str_pad_start" => fn_call!(self, "ecma:string", "padStart", 3),
            "str_pad_end" => fn_call!(self, "ecma:string", "padEnd", 3),
            "str_compare" => fn_call!(self, "wasm:js-string", "compare", 2),
            "str_concat" => fn_call!(self, "wasm:js-string", "concat", 2),
            // Array primitives — every emit flows through
            // `common::collections::*` so the emitted bytecode uses
            // `ecma:array.*` imports. One-place-to-change: flip the
            // provider in collections.rs and every array op in every
            // language re-routes.
            "array_push" => {
                let l = self.line;
                common::collections::emit_push(&mut self.chunks, self.current, l);
            }
            "array_pop" => {
                let l = self.line;
                common::collections::emit_pop(&mut self.chunks, self.current, l);
            }
            "array_shift" => {
                let l = self.line;
                common::collections::emit_shift(&mut self.chunks, self.current, l);
            }
            "array_reverse" => {
                let l = self.line;
                common::collections::emit_reverse(&mut self.chunks, self.current, l);
            }
            "array_join" => {
                let l = self.line;
                common::collections::emit_join(&mut self.chunks, self.current, l);
            }
            "array_concat" => {
                let l = self.line;
                common::collections::emit_concat(&mut self.chunks, self.current, l);
            }
            "array_fill" => {
                let l = self.line;
                common::collections::emit_fill(&mut self.chunks, self.current, l);
            }
            "array_length" => {
                let l = self.line;
                common::collections::emit_len(&mut self.chunks, self.current, l);
            }
            "array_slice" => {
                let l = self.line;
                common::collections::emit_slice(&mut self.chunks, self.current, l);
            }
            // The dynamic languages' lenient subscript is `emit_get`/`emit_set`
            // (JS/Python/PHP return undefined / no-op out of bounds). A spec WASM
            // profile (`function_references`, i.e. wast) uses the GC `array.get`/
            // `array.set` that TRAP on a null array or out-of-bounds index.
            "array_get" => {
                let l = self.line;
                if self.profile.function_references {
                    common::collections::emit_gc_array_get(&mut self.chunks, self.current, l);
                } else {
                    common::collections::emit_get(&mut self.chunks, self.current, l);
                }
            }
            "array_set" => {
                let l = self.line;
                if self.profile.function_references {
                    common::collections::emit_gc_array_set(&mut self.chunks, self.current, l);
                } else {
                    common::collections::emit_set(&mut self.chunks, self.current, l);
                }
            }
            "array_contains" => {
                let l = self.line;
                common::collections::emit_contains(&mut self.chunks, self.current, l);
            }
            "array_index_of" => {
                let l = self.line;
                common::collections::emit_index_of(&mut self.chunks, self.current, l);
            }
            _ => {
                // Single source of truth: resolve any remaining opcode straight
                // from the VM's opcode table rather than maintaining a second
                // hardcoded list. Profile names are underscore-form
                // (`i32_clz`); WASM mnemonics dot the type prefix (`i32.clz`).
                if let Some(op) = Op::from_flattened_name(op_name) {
                    self.emit(op);
                } else {
                    let c = self.str_const(op_name);
                    self.emit_u16(Op::GLOBAL_GET, c);
                }
            }
        }
    }

    pub(super) fn emit_builtin_opcode(
        &mut self,
        op_name: &str,
        args: &[&Expression],
    ) -> Result<(), String> {
        let line = self.line;
        match op_name {
            "abs" => {
                self.compile_expr(args[0])?;
                common::math::emit_abs(self.chunk(), line);
            }
            "sqrt" => {
                self.compile_expr(args[0])?;
                common::math::emit_sqrt(self.chunk(), line);
            }
            "round" => {
                if args.len() >= 2 {
                    let number = self.import("ecma:number", "Number");
                    let scale_slot = self.define_local("__round_scale");
                    self.emit_const(Value::F64(10.0));
                    self.compile_expr(args[1])?;
                    common::math::emit_pow(self.chunk(), line);
                    self.emit_host_call(number, 1);
                    self.emit_u16(Op::LOCAL_SET, scale_slot);

                    self.compile_expr(args[0])?;
                    self.emit_host_call(number, 1);
                    self.emit_u16(Op::LOCAL_GET, scale_slot);
                    self.emit(Op::F64_MUL);
                    common::math::emit_round(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, scale_slot);
                    self.emit(Op::F64_DIV);
                } else {
                    self.compile_expr(args[0])?;
                    common::math::emit_round(self.chunk(), line);
                }
            }
            "trunc" => {
                self.compile_expr(args[0])?;
                common::math::emit_trunc(self.chunk(), line);
            }
            "floor" => {
                self.compile_expr(args[0])?;
                common::math::emit_floor(self.chunk(), line);
            }
            "ceil" => {
                self.compile_expr(args[0])?;
                common::math::emit_ceil(self.chunk(), line);
            }
            "min" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::math::emit_min(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "max" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::math::emit_max(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "sqr" => {
                self.compile_expr(args[0])?;
                inst!(self, core_wasm::dup);
                self.emit(Op::F64_MUL);
            }
            "succ" => {
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(1.0));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                };
            }
            "pred" => {
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(1.0));
                self.emit(Op::F64_SUB);
            }
            "to_upper" => {
                self.compile_expr(args[0])?;
                common::strings::emit_to_upper(self.chunk(), line);
            }
            "to_lower" => {
                self.compile_expr(args[0])?;
                common::strings::emit_to_lower(self.chunk(), line);
            }
            "trim" => {
                self.compile_expr(args[0])?;
                common::strings::emit_trim(self.chunk(), line);
            }
            "str_contains" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    fn_call!(self, "ecma:string", "includes", 2);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "str_starts_with" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    fn_call!(self, "ecma:string", "startsWith", 2);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "str_ends_with" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    fn_call!(self, "ecma:string", "endsWith", 2);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "concat" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                common::strings::emit_concat(self.chunk(), args.len(), line);
            }
            "replace" => {
                if args.len() >= 3 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[2])?;
                    common::strings::emit_replace(self.chunk(), line);
                }
            }
            "repeat" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::strings::emit_repeat(self.chunk(), line);
                }
            }
            "leftstr" => {
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(0.0));
                self.compile_expr(args[1])?;
                common::strings::emit_substring(self.chunk(), line);
            }
            "high" => {
                self.compile_expr(args[0])?;
                common::strings::emit_length(self.chunk(), line);
                self.emit_const(Value::F64(1.0));
                self.emit(Op::F64_SUB);
            }
            "low" => {
                self.emit_const(Value::F64(0.0));
            }
            "setlength" => {
                if args.len() >= 2 {
                    self.compile_setlength(args[0], args[1])?;
                } else {
                    self.emit(Op::NULL);
                }
            }
            "trim_start" => {
                self.compile_expr(args[0])?;
                common::strings::emit_trim_start(self.chunk(), line);
            }
            "trim_end" => {
                self.compile_expr(args[0])?;
                common::strings::emit_trim_end(self.chunk(), line);
            }
            "pow" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::math::emit_pow(self.chunk(), line);
                }
            }
            "log" => {
                self.compile_expr(args[0])?;
                common::math::emit_log(self.chunk(), line);
            }
            "sin" => {
                self.compile_expr(args[0])?;
                common::math::emit_sin(self.chunk(), line);
            }
            "cos" => {
                self.compile_expr(args[0])?;
                common::math::emit_cos(self.chunk(), line);
            }
            "tan" => {
                self.compile_expr(args[0])?;
                common::math::emit_tan(self.chunk(), line);
            }
            "exp" => {
                self.compile_expr(args[0])?;
                common::math::emit_exp(self.chunk(), line);
            }
            "is_null" => {
                self.compile_expr(args[0])?;
                self.emit(Op::REF_IS_NULL);
            }
            "space" => {
                self.emit_const(Value::String(Arc::from(" ")));
                self.compile_expr(args[0])?;
                common::strings::emit_repeat(self.chunk(), line);
            }
            "assigned" => {
                self.compile_expr(args[0])?;
                self.emit(Op::REF_IS_NULL);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            }
            "freeandnil" => {
                if let Some(first) = args.first() {
                    if let ExprKind::Ident(var) = &first.kind {
                        let var = var.clone();
                        self.emit(Op::NULL);
                        self.emit_var_set(&var);
                    }
                }
                self.emit(Op::NULL);
            }
            // NOTE: bare WASM opcodes (f64.*, ref.is_null, ALL SIMD lane ops, …)
            // are NOT enumerated here — they emit through the one VM opcode table
            // via `Op::from_wasm_name` in the `_ =>` fallback. Only genuine recipes
            // (host calls, multi-op sequences) and VM-internal conversions stay.
            "i32_from_f64" | "to_int" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                let line = self.line;
                common::convert::emit_to_int(self.chunk(), line);
            }
            "f64_from_i32" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                self.emit(Op::F64_FROM_I32);
            }
            "dyn_to_bool" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                {
                    let line = self.line;
                    if self.is_python_profile() {
                        self.emit_condition_truthiness_from_stack();
                    } else {
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    }
                }
            }
            "ref_is_array" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                fn_call!(self, "ecma:array", "isArray", 1);
            }
            "ref_typeof" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                fn_call!(self, "ecma:value", "typeof", 1);
            }
            "str_length" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                fn_call!(self, "wasm:js-string", "length", 1);
            }
            "str_to_upper" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                fn_call!(self, "ecma:string", "toUpperCase", 1);
            }
            "str_to_lower" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                fn_call!(self, "ecma:string", "toLowerCase", 1);
            }
            "str_trim" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                fn_call!(self, "ecma:string", "trim", 1);
            }
            "str_trim_start" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                fn_call!(self, "ecma:string", "trimStart", 1);
            }
            "str_trim_end" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                fn_call!(self, "ecma:string", "trimEnd", 1);
            }
            "str_reverse" => {
                for a in args {
                    self.compile_expr(a)?;
                }
                {
                    let l = self.line;
                    crate::primitives::strings::emit_str_reverse(self.chunk(), l)
                };
            }
            "str_last_index_of" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    fn_call!(self, "ecma:string", "lastIndexOf", 2);
                }
            }
            "str_from_char_code" => {
                // String.fromCharCode(72, 105) → "Hi"
                self.compile_expr(args[0])?;
                common::convert::emit_to_int(self.chunk(), line);
                fn_call!(self, "wasm:js-string", "fromCharCode", 1);
                for a in &args[1..] {
                    self.compile_expr(a)?;
                    common::convert::emit_to_int(self.chunk(), line);
                    fn_call!(self, "wasm:js-string", "fromCharCode", 1);
                    fn_call!(self, "wasm:js-string", "concat", 2);
                }
            }
            "str_compare" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    fn_call!(self, "wasm:js-string", "compare", 2);
                }
            }
            "str_split" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    fn_call!(self, "ecma:string", "split", 2);
                }
            }
            "str_getcsv" => {
                for arg in args {
                    self.compile_expr(arg)?;
                }
                vybe_runtime::registry::hooks(&self.profile.name)
                    .str_getcsv
                    .unwrap()(
                    &mut self.chunks, self.current, args.len() as u8, line
                );
            }
            "str_repeat" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    fn_call!(self, "ecma:string", "repeat", 2);
                }
            }
            "array_join" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    {
                        let l = self.line;
                        common::collections::emit_join(&mut self.chunks, self.current, l);
                    }
                }
            }
            "set_timer" => {
                if let Some(cb) = args.first() {
                    self.compile_expr(cb)?;
                } else {
                    self.emit(Op::NULL);
                }
                if args.len() >= 2 {
                    self.compile_expr(args[1])?;
                } else {
                    inst!(self, core_wasm::i32_const, 0);
                }
                fn_call!(self, "web:timers", "setTimeout", 2);
            }
            // Array primitives — every caller dispatches through
            // `common::collections::*`, which now routes to `ecma:array.*`
            // imports (Phase D). Keep the arg-evaluation and stack shape
            // details here; the emit itself lives in compiler_common so
            // the identical surface is used by every language.
            "array_length" => {
                if let Some(first) = args.first() {
                    self.compile_expr(first)?;
                    common::collections::emit_len(&mut self.chunks, self.current, line);
                } else {
                    self.emit_const(Value::I32(0));
                }
            }
            "array_push" => {
                // PHP `array_push($a, v1, v2, ...)` — push each value.
                // Returns the new length (of the last push).
                if let Some(arr) = args.first() {
                    if args.len() == 1 {
                        self.compile_expr(arr)?;
                        common::collections::emit_len(&mut self.chunks, self.current, line);
                    } else {
                        let tail = args.len() - 1;
                        for (i, v) in args[1..].iter().enumerate() {
                            self.compile_expr(arr)?;
                            self.compile_expr(v)?;
                            common::collections::emit_push(&mut self.chunks, self.current, line);
                            // Drop intermediate lengths; the final one
                            // is the expression's value.
                            if i != tail - 1 {
                                self.emit(Op::DROP);
                            }
                        }
                    }
                } else {
                    self.emit_const(Value::I32(0));
                }
            }
            "array_pop" => {
                if let Some(first) = args.first() {
                    self.compile_expr(first)?;
                    common::collections::emit_pop(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "array_shift" => {
                if let Some(first) = args.first() {
                    self.compile_expr(first)?;
                    common::collections::emit_shift(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "array_reverse" => {
                if let Some(first) = args.first() {
                    self.compile_expr(first)?;
                    common::collections::emit_reverse(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "array_concat" => {
                if args.is_empty() {
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                } else {
                    self.compile_expr(args[0])?;
                    for v in &args[1..] {
                        self.compile_expr(v)?;
                        common::collections::emit_concat(&mut self.chunks, self.current, line);
                    }
                }
            }
            "array_index_of" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::collections::emit_index_of(&mut self.chunks, self.current, line);
                } else {
                    self.emit_const(Value::I32(-1));
                }
            }
            // PHP `in_array($needle, $haystack)` — walker already normalized
            // arg order to [haystack, needle, strict?] matching JS's
            // `arr.includes(needle, fromIndex?)`. emit_contains calls
            // `ecma:array.includes` which is polymorphic over Array,
            // Map, and Ordinary, so PHP's `in_array` works uniformly on
            // assoc arrays, indexed arrays, and superglobals.
            "array_contains" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::collections::emit_contains(&mut self.chunks, self.current, line);
                } else {
                    inst!(self, core_wasm::bool_const, false);
                }
            }
            _ => {
                // Multi-memory selector suffix from the wast walker
                // (`i32.store@@mem1`, or `memory.copy@@mem<dst>@@mem<src>` with
                // two indices): non-default linear memories. Strip the suffixes
                // to resolve the base opcode; each selected memidx is emitted
                // after the opcode as the VM's fixed 4-byte `0xEE 0x00 <u16>`
                // selector, in order (one per positional memidx the VM reads).
                let (op_name, mem_selectors): (&str, Vec<u32>) = {
                    let mut parts = op_name.split("@@mem");
                    let base = parts.next().unwrap_or(op_name);
                    let sels: Vec<u32> = parts.filter_map(|p| p.parse::<u32>().ok()).collect();
                    (base, sels)
                };
                // Single source of truth: resolve any remaining opcode straight
                // from the VM's opcode table. The VM's `operand_format` (same
                // table) tells us how to encode any immediates — lane index,
                // v128.const value, shuffle mask — so there is no second list.
                let resolved = Op::from_flattened_name(op_name);
                let Some(op) = resolved else {
                    self.emit(Op::NULL);
                    return Ok(());
                };
                use vybe_runtime::opcode::OperandFormat;
                let l = self.line;
                match op.operand_format() {
                    // v128.const: args are all immediates — a shape token then
                    // the lane values — encoded to the 16-byte vector.
                    OperandFormat::V128Const => {
                        let bytes = encode_v128_const(args);
                        self.chunk().emit_op(Op::V128_CONST, l);
                        for b in bytes {
                            self.chunk().emit(b, l);
                        }
                    }
                    // Type/index immediate (array.new $t / array.get_s $t / …):
                    // fold puts the immediate first, then the stack operands.
                    OperandFormat::U16 => {
                        for a in &args[1..] {
                            self.compile_expr(a)?;
                        }
                        let imm = expr_const_u16(args.first().copied());
                        self.chunk().emit_op_u16(op, imm, l);
                    }
                    // Two index immediates then stack operands: the GC
                    // array-from-segment ops (`array.new_data $T $d`,
                    // `array.new_elem`, `array.init_data`, `array.init_elem`)
                    // carry `typeidx` + `dataidx`/`elemidx`. Without this arm they
                    // fell to the plain `_` arm, emitted NO immediates, and the VM
                    // read the next opcode's bytes as the segment index.
                    OperandFormat::U16_U16 => {
                        for a in args.iter().skip(2) {
                            self.compile_expr(a)?;
                        }
                        let imm1 = expr_const_u16(args.first().copied());
                        let imm2 = expr_const_u16(args.get(1).copied());
                        self.chunk().emit_op_u16(op, imm1, l);
                        self.chunk().emit((imm2 >> 8) as u8, l);
                        self.chunk().emit((imm2 & 0xff) as u8, l);
                    }
                    // Lane ops (extract_lane / replace_lane): the fold puts the
                    // lane immediate first, then the stack operands.
                    OperandFormat::U8 => {
                        for a in &args[1..] {
                            self.compile_expr(a)?;
                        }
                        let lane = expr_const_u8(args.first().copied());
                        self.chunk().emit_op(op, l);
                        self.chunk().emit(lane, l);
                    }
                    // SIMD lane memory ops (`v128.load8_lane` / `v128.store32_lane`
                    // / …): an optional memarg then a lane byte. The wast fold
                    // supplies just the lane index (memarg defaults to 0), so we
                    // emit align=0 (the VM reads that as "no offset/memidx") then
                    // the lane byte. Stack operands (addr, vector) follow.
                    OperandFormat::MemLane => {
                        // SIMD lane mem op. The VM pops the top operand first:
                        // `load*_lane` wants the vector on top (`[addr vector]`),
                        // `store*_lane` wants the address on top (`[vector addr]`).
                        // The fold hands operands in source (deepest-first) order,
                        // so push them reversed to land the right one on top.
                        // Only a lane byte is emitted — the VM's optional-memarg
                        // peek never consumes a byte because lane indices are
                        // < 0x80 (so the byte reads back as the lane).
                        for a in args[1..].iter().rev() {
                            self.compile_expr(a)?;
                        }
                        let lane = expr_const_u8(args.first().copied());
                        self.chunk().emit_op(op, l);
                        self.chunk().emit(lane, l);
                    }
                    // Two byte immediates then stack operands (call_indirect:
                    // argc, tableidx). The fold puts both immediates first.
                    OperandFormat::U8_U8 => {
                        for a in &args[2..] {
                            self.compile_expr(a)?;
                        }
                        self.chunk().emit_op(op, l);
                        self.chunk().emit(expr_const_u8(args.first().copied()), l);
                        self.chunk().emit(expr_const_u8(args.get(1).copied()), l);
                    }
                    // Three byte immediates then stack operands (call_indirect:
                    // argc, tableidx, expected result count). The fold puts all
                    // three immediates first.
                    OperandFormat::U8_U8_U8 => {
                        for a in &args[3..] {
                            self.compile_expr(a)?;
                        }
                        self.chunk().emit_op(op, l);
                        self.chunk().emit(expr_const_u8(args.first().copied()), l);
                        self.chunk().emit(expr_const_u8(args.get(1).copied()), l);
                        self.chunk().emit(expr_const_u8(args.get(2).copied()), l);
                    }
                    // i8x16.shuffle: 16 lane-index immediates, then two vectors.
                    OperandFormat::Shuffle => {
                        for a in args.iter().skip(16) {
                            self.compile_expr(a)?;
                        }
                        self.chunk().emit_op(op, l);
                        for i in 0..16 {
                            self.chunk().emit(expr_const_u8(args.get(i).copied()), l);
                        }
                    }
                    // Plain opcode: operands on the stack, no immediate.
                    _ => {
                        for a in args {
                            self.compile_expr(a)?;
                        }
                        self.emit(op);
                        // Multi-memory selectors, read by the VM's
                        // `read_optional_memarg`/`read_optional_memidx_immediate`.
                        // VM instructions are always 4 bytes, so each selector is a
                        // fixed 4-byte block (`0xEE 0x00 <memidx u16 BE>`) that keeps
                        // the following instruction 4-aligned. `memory.copy` emits
                        // two (dst then src); load/store/size/grow/fill emit one.
                        for midx in &mem_selectors {
                            self.chunk().emit(0xEE, l);
                            self.chunk().emit(0x00, l);
                            self.chunk().emit((midx >> 8) as u8, l);
                            self.chunk().emit((midx & 0xff) as u8, l);
                        }
                    }
                }
            }
        }
        Ok(())
    }

    /// Emit a multi-opcode intrinsic sequence.
    pub(super) fn emit_fortran_scan_like(
        &mut self,
        args: &[&Expression],
        invert_match: bool,
    ) -> Result<(), String> {
        let line = self.line;
        if args.len() < 2 {
            self.emit(Op::NULL);
            return Ok(());
        }

        let source_slot = self.define_local("__fortran_scan_source");
        let set_slot = self.define_local("__fortran_scan_set");
        let back_slot = self.define_local("__fortran_scan_back");
        let len_slot = self.define_local("__fortran_scan_len");
        let index_slot = self.define_local("__fortran_scan_index");
        let result_slot = self.define_local("__fortran_scan_result");

        self.compile_expr(args[0])?;
        self.emit_u16(Op::LOCAL_SET, source_slot);

        self.compile_expr(args[1])?;
        self.emit_u16(Op::LOCAL_SET, set_slot);

        if let Some(back_arg) = args.get(2) {
            self.compile_expr(back_arg)?;
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            };
        } else {
            inst!(self, core_wasm::bool_const, false);
        }
        self.emit_u16(Op::LOCAL_SET, back_slot);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        common::strings::emit_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, len_slot);

        inst!(self, core_wasm::i32_const, 0);
        self.emit_u16(Op::LOCAL_SET, result_slot);

        self.emit_u16(Op::LOCAL_GET, back_slot);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, len_slot);
        inst!(self, core_wasm::i32_const, 1);
        self.emit(Op::I32_SUB);
        self.emit_u16(Op::LOCAL_SET, index_slot);

        let back_block = self.chunk().emit_block(line);
        let (back_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        inst!(self, core_wasm::i32_const, 0);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
        };
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, set_slot);
        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        fn_call!(self, "ecma:string", "charAt", 2);
        fn_call!(self, "ecma:string", "includes", 2);
        if invert_match {
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_not(self.chunk(), line);
            };
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::I32(1));
        self.emit(Op::I32_ADD);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        // depth 0=inner IF, depth 1=back_loop (LOOP→repeats), depth 2=back_block (BLOCK→exits)
        self.chunk().emit_br(2, line);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        inst!(self, core_wasm::i32_const, 1);
        self.emit(Op::I32_SUB);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(back_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(back_block);

        self.chunk().emit_else(line);
        inst!(self, core_wasm::i32_const, 0);
        self.emit_u16(Op::LOCAL_SET, index_slot);

        let forward_block = self.chunk().emit_block(line);
        let (forward_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
        };
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, set_slot);
        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        fn_call!(self, "ecma:string", "charAt", 2);
        fn_call!(self, "ecma:string", "includes", 2);
        if invert_match {
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_not(self.chunk(), line);
            };
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::I32(1));
        self.emit(Op::I32_ADD);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        // depth 0=inner IF, depth 1=forward_loop (LOOP→repeats), depth 2=forward_block (BLOCK→exits)
        self.chunk().emit_br(2, line);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        inst!(self, core_wasm::i32_const, 1);
        self.emit(Op::I32_ADD);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(forward_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(forward_block);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(())
    }

    pub(super) fn emit_intrinsic(
        &mut self,
        name: &str,
        args: &[&Expression],
    ) -> Result<(), String> {
        let line = self.line;
        match name {
            "cstr" => {
                self.compile_expr(args[0])?;
                let value_slot = self.define_local("__vb_cstr_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);

                self.emit_u16(Op::LOCAL_GET, value_slot);
                fn_call!(self, "ecma:value", "typeof", 1);
                self.emit_const(Value::String(Arc::from("boolean")));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);

                self.emit_u16(Op::LOCAL_GET, value_slot);
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                self.emit_const(Value::String(Arc::from("True")));
                self.chunk().emit_else(line);
                self.emit_const(Value::String(Arc::from("False")));
                self.chunk().emit_end(line);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                fn_call!(self, "ecma:value", "typeof", 1);
                self.emit_const(Value::String(Arc::from("object")));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);

                self.emit_u16(Op::LOCAL_GET, value_slot);
                let type_key = self
                    .chunk()
                    .add_constant(Value::String(Arc::from("__type")));
                self.emit_u16(Op::STRUCT_GET, type_key);
                self.emit_const(Value::String(Arc::from("DateTime")));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let time_key = self
                    .chunk()
                    .add_constant(Value::String(Arc::from("__time")));
                self.emit_u16(Op::STRUCT_GET, time_key);
                let iso_idx = self.import("ecma:date", "toISOString");
                self.emit_host_call(iso_idx, 1);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let string_idx = self.import("ecma:string", "String");
                self.emit_host_call(string_idx, 1);
                self.chunk().emit_end(line);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let string_idx = self.import("ecma:string", "String");
                self.emit_host_call(string_idx, 1);
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
            }
            "cbyte" => {
                self.compile_expr(args[0])?;
                common::convert::emit_to_int(self.chunk(), line);
                self.emit_const(Value::I32(0xFF));
                self.emit(Op::I32_AND);
            }
            "ubound" => {
                self.compile_expr(args[0])?;
                common::collections::emit_len(&mut self.chunks, self.current, line);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_SUB);
            }
            "lbound" => {
                inst!(self, core_wasm::i32_const, 0);
            }
            "erase" => {
                // VB `Erase arr` — releases / clears the array contents. For
                // dynamic arrays, real VB frees the storage and leaves the
                // variable referring to an uninitialised array; for
                // fixed-size arrays, it re-zeros each element. We return a
                // fresh empty array, which satisfies both reads (`.Length`
                // works, yields 0) and assignment (`arr = Erase(arr)`).
                //
                // The arg is still compiled for any side effects and then
                // dropped — matches the VB semantic that the old binding is
                // released.
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    self.emit(Op::DROP);
                }
                let l = self.line;
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, l);
            }
            "readline" => {
                // wasi:cli/stdin.get-stdin → [method]input-stream.blocking-read
                crate::primitives::io::emit_input(self.chunk(), line);
            }
            "write_stdout" => {
                // libc stdout write → wasi:io DIRECTLY (no `print`/wasi:logging,
                // no vybelib). arg0 = exact bytes; byte-faithful, no implicit
                // newline. Mirrors the proven wasi:cli/stdout.get-stdout +
                // wasi:io/streams.blocking-write-and-flush path.
                self.compile_expr(args[0])?;
                let text_slot = self.define_local("__c_wasi_stdout_text");
                self.emit_u16(Op::LOCAL_SET, text_slot);
                let stdout_idx = self.import("wasi:cli/stdout", "get-stdout");
                let write_idx = self.import(
                    "wasi:io/streams",
                    "[method]output-stream.blocking-write-and-flush",
                );
                self.emit_host_call(stdout_idx, 0);
                self.emit_u16(Op::LOCAL_GET, text_slot);
                self.emit_host_call(write_idx, 2);
                self.emit(Op::DROP);
                // fputs/stdout_append return 0
                self.emit_const(Value::I32(0));
            }
            "asc" => {
                self.compile_expr(args[0])?;
                inst!(self, core_wasm::i32_const, 0);
                fn_call!(self, "wasm:js-string", "charCodeAt", 2);
            }
            "space" => {
                self.emit_const(Value::String(Arc::from(" ")));
                self.compile_expr(args[0])?;
                common::convert::emit_to_int(self.chunk(), line);
                common::strings::emit_repeat(self.chunk(), line);
            }
            "isobject" => {
                if let Some(arg) = args.first() {
                    if let Some(result) = self.vb_is_object_expr(arg) {
                        self.emit_const(Value::Bool(result));
                    } else {
                        self.compile_expr(arg)?;
                        let value_slot = self.define_local("__vb_isobject_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        fn_call!(self, "ecma:array", "isArray", 1);
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if_value(line);
                        self.emit_const(Value::Bool(true));
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        inst!(self, recipes::is_object);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_end(line);
                    }
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "isreference" => {
                if let Some(arg) = args.first() {
                    if let Some(result) = self.vb_is_reference_expr(arg) {
                        self.emit_const(Value::Bool(result));
                    } else {
                        self.compile_expr(arg)?;
                        let value_slot = self.define_local("__vb_isref_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot);

                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        fn_call!(self, "wasm:js-string", "test", 1);
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if_value(line);
                        self.emit_const(Value::Bool(true));
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        fn_call!(self, "ecma:array", "isArray", 1);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if_value(line);
                        self.emit_const(Value::Bool(true));
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        inst!(self, recipes::is_object);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                    }
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "typename" => {
                if let Some(arg) = args.first() {
                    if let Some(name) = self.vb_typename_from_expr(arg) {
                        self.emit_const(Value::String(Arc::from(name)));
                    } else {
                        self.compile_expr(arg)?;
                        fn_call!(self, "ecma:value", "typeof", 1);
                    }
                } else {
                    self.emit_const(Value::String(Arc::from("Nothing")));
                }
            }
            "command" => {
                let args_idx = self.import("wasi:cli/environment", "get-arguments");
                self.emit_host_call(args_idx, 0);
                self.emit_const(Value::String(Arc::from(" ")));
                common::collections::emit_join(&mut self.chunks, self.current, line);
            }
            "environ" => {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let env_idx = self.import("wasi:cli/environment", "get-environment");
                    self.emit_host_call(env_idx, 1);
                } else {
                    self.emit_const(Value::String(Arc::from("")));
                }
            }
            "timer" => {
                // Timer = seconds since midnight.
                // ecma:date.now() → ms, then extract UTC H/M/S → h*3600+m*60+s
                let now_idx = self.import("ecma:date", "now");
                let get_h_idx = self.import("ecma:date", "getUTCHours");
                let get_m_idx = self.import("ecma:date", "getUTCMinutes");
                let get_s_idx = self.import("ecma:date", "getUTCSeconds");
                self.emit_host_call(now_idx, 0);
                let ms_slot = self.define_local("__vb_timer_ms");
                self.emit_u16(Op::LOCAL_SET, ms_slot);
                // hours
                self.emit_u16(Op::LOCAL_GET, ms_slot);
                self.emit_host_call(get_h_idx, 1);
                self.emit_const(Value::F64(3600.0));
                self.emit(Op::F64_MUL);
                // + minutes * 60
                self.emit_u16(Op::LOCAL_GET, ms_slot);
                self.emit_host_call(get_m_idx, 1);
                self.emit_const(Value::F64(60.0));
                self.emit(Op::F64_MUL);
                self.emit(Op::F64_ADD);
                // + seconds
                self.emit_u16(Op::LOCAL_GET, ms_slot);
                self.emit_host_call(get_s_idx, 1);
                self.emit(Op::F64_ADD);
            }
            "switch" => {
                if args.len() < 2 {
                    self.emit(Op::NULL);
                } else {
                    let mut slots = Vec::with_capacity(args.len());
                    for (index, arg) in args.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let slot = self.define_local(&format!("__vb_switch_{index}"));
                        self.emit_u16(Op::LOCAL_SET, slot);
                        slots.push(slot);
                    }

                    let result_slot = self.define_local("__vb_switch_result");
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    let matched_slot = self.define_local("__vb_switch_matched");
                    self.emit_const(Value::Bool(false));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    for pair in slots.chunks(2) {
                        if pair.len() < 2 {
                            break;
                        }
                        self.emit_u16(Op::LOCAL_GET, matched_slot);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.emit(Op::I32_EQZ);
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, pair[0]);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, pair[1]);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit_const(Value::Bool(true));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                    }
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                }
            }
            "string_repeat" => {
                // String(n, char): VB arg order reversed
                if args.len() >= 2 {
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::strings::emit_repeat(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "left" => {
                // Left(s, n) → substring(s, 0, n)
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    inst!(self, core_wasm::i32_const, 0);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "php_is_float" => {
                // PHP `is_float` — true only for non-integer numbers.
                // Composes ecma:number.isInteger + boolean negation
                // with a leading `typeof v === "number"` guard so
                // strings / objects don't match (REF_IS_NUMBER opcode
                // covers the typeof-number predicate).
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    let v_slot = self.define_local("__php_isf_v");
                    self.emit_u16(Op::LOCAL_SET, v_slot);
                    self.emit_u16(Op::LOCAL_GET, v_slot);
                    fn_call!(self, "wasm:js-number", "test", 1);
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, v_slot);
                    let is_int_idx = self.import("ecma:number", "isInteger");
                    self.emit_host_call(is_int_idx, 1);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    };
                    self.chunk().emit_else(line);
                    self.emit_const(Value::Bool(false));
                    self.chunk().emit_end(line);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_string" => {
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    fn_call!(self, "wasm:js-string", "test", 1);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_array" => {
                // PHP `is_array` matches any of: ObjectKind::Array,
                // ObjectKind::Map, ObjectKind::Ordinary (plain assoc
                // object). REF_IS_ARRAY only checks Array; we layer
                // an Object check via REF_IS_OBJECT (covers Map and
                // Ordinary too — both are Object-kind values).
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    inst!(self, recipes::is_object);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_bool" => {
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    fn_call!(self, "wasm:js-boolean", "test", 1);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_null" => {
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    self.emit(Op::REF_IS_NULL);
                } else {
                    self.emit_const(Value::Bool(true));
                }
            }
            "php_is_object" => {
                // PHP `is_object` matches user objects but NOT plain
                // arrays. Approximated as REF_IS_OBJECT && !is_array.
                // For Phase-1 simplicity the same predicate as is_array
                // — distinction requires a class-instance vs assoc-array
                // tag which Vybe doesn't track yet.
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    inst!(self, recipes::is_object);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_defined" => {
                if let Some(Expression {
                    kind: ExprKind::Lit(Literal::Str(name)),
                    ..
                }) = args.first()
                {
                    let global_name = self.canon(name);
                    if is_php_builtin_constant_name(&global_name) {
                        self.emit_const(Value::Bool(true));
                        return Ok(());
                    }
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                    fn_call!(self, "ecma:value", "typeof", 1);
                    self.emit_const(Value::String(Arc::from("undefined")));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    };
                } else {
                    if let Some(arg) = args.first() {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    self.emit_const(Value::Bool(false));
                }
            }
            // `constant("NAME")` — read back the global that `define` wrote.
            // Only the literal-name form is compilable to a direct global
            // read; a dynamic name yields NULL (no runtime global-by-name
            // surface yet).
            "php_constant" => {
                if let Some(Expression {
                    kind: ExprKind::Lit(Literal::Str(name)),
                    ..
                }) = args.first()
                {
                    let global_name = self.canon(name);
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                } else {
                    if let Some(arg) = args.first() {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    self.emit(Op::NULL);
                }
            }
            "php_function_exists" => {
                if let Some(Expression {
                    kind: ExprKind::Lit(Literal::Str(name)),
                    ..
                }) = args.first()
                {
                    let builtin_exists = self.profile.lookup_builtin(name).is_some()
                        || crate::primitives::imports::resolve_common_import(name).is_some();
                    if builtin_exists {
                        self.emit_const(Value::Bool(true));
                    } else {
                        let lowered = name.to_ascii_lowercase();
                        let exists = self.defined_functions.iter().any(|function_name| {
                            Self::php_function_name_lookup_spellings(function_name)
                                .iter()
                                .any(|spelling| spelling == &lowered)
                        });
                        self.emit_const(Value::Bool(exists));
                    }
                } else {
                    if let Some(arg) = args.first() {
                        self.compile_expr(arg)?;
                        let name_slot = self.define_local("__php_function_exists_name");
                        self.emit_u16(Op::LOCAL_SET, name_slot);

                        self.emit_u16(Op::LOCAL_GET, name_slot);
                        fn_call!(self, "ecma:value", "typeof", 1);
                        self.emit_const(Value::String(Arc::from("string")));
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                        };
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_if_value(line);

                        self.emit_u16(Op::LOCAL_GET, name_slot);
                        fn_call!(self, "ecma:string", "toLowerCase", 1);
                        let lowered_slot = self.define_local("__php_function_exists_lowered");
                        self.emit_u16(Op::LOCAL_SET, lowered_slot);

                        let mut known_functions: Vec<String> =
                            self.defined_functions.iter().cloned().collect();
                        known_functions.sort();
                        let exists_slot = self.define_local("__php_function_exists_result");
                        self.emit_const(Value::Bool(false));
                        self.emit_u16(Op::LOCAL_SET, exists_slot);
                        for function_name in known_functions {
                            for spelling in Self::php_function_name_lookup_spellings(&function_name)
                            {
                                self.emit_u16(Op::LOCAL_GET, lowered_slot);
                                self.emit_const(Value::String(Arc::from(spelling)));
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                                };
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                                };
                                self.chunk().emit_if(line);
                                self.emit_const(Value::Bool(true));
                                self.emit_u16(Op::LOCAL_SET, exists_slot);
                                self.chunk().emit_end(line);
                            }
                        }
                        self.emit_u16(Op::LOCAL_GET, exists_slot);
                        self.chunk().emit_else(line);
                        self.emit_const(Value::Bool(false));
                        self.chunk().emit_end(line);
                    } else {
                        self.emit_const(Value::Bool(false));
                    }
                }
            }
            // `symbol_exists` / `symbol_exists:<kind>` — the one primitive behind
            // `class_exists` / `interface_exists` / `trait_exists` /
            // `enum_exists` / `kind_of?`. They differ only in which declared
            // kind they accept, so the kind is a parameter from the profile's
            // builtin table rather than four near-identical intrinsics.
            //
            // The answer comes from the `__kind` annotation the class compiler
            // stamps, so it is true for a type defined after compilation — by an
            // autoloader, by `eval`, or in another file of the same bundle.
            // A second argument of literal `false` suppresses resolver
            // consultation (a language's "autoload" flag).
            // `symbol_resolve_or_throw:<ExceptionName>` — resolve a declared
            // type by name at runtime and yield it, or throw when it does not
            // resolve. Java `Class.forName` (ClassNotFoundException), and the
            // same shape serves Python/Ruby `NameError`. The exception's
            // spelling is profile data; the mechanism is shared.
            //
            // Resolution goes through the same constructor-global-ref path as
            // the `*_exists` family, so a language with a registered resolver
            // stack (PHP autoload, a Java ClassLoader) gets a chance to supply
            // the type before this decides it is missing.
            // `symbol_probe` — resolve a declared type by name and yield its
            // NAME, or `null` when it does not resolve. .NET
            // `Type.GetType(name)` returns null on a miss rather than throwing
            // (it only throws when asked, `throwOnError: true`), so this is the
            // resolve-or-null counterpart of `symbol_require`.
            //
            // Yields the name because .NET, like Java, represents a type BY its
            // name: `typeof(X).Name` is `"X"`, so `Type.GetType("X")` has to
            // agree with `typeof(X)` rather than invent a second representation.
            "symbol_probe" => {
                if let Some(Expression {
                    kind: ExprKind::Lit(Literal::Str(name)),
                    ..
                }) = args.first()
                {
                    for arg in args.iter().skip(1) {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    let global_name = self.canon_type_global(name);
                    self.emit_constructor_global_ref(&global_name, name);
                    let line = self.line;
                    let resolved = self.define_local("__symbol_probe");
                    self.emit_u16(Op::LOCAL_SET, resolved);
                    self.emit_u16(Op::LOCAL_GET, resolved);
                    {
                        let idx = self.chunk().add_import("wasm:js-undefined", "test");
                        self.chunk().emit_call(idx, 1, line);
                    }
                    self.chunk().emit_if_value(line);
                    self.emit(Op::NULL);
                    self.chunk().emit_else(line);
                    self.emit_const(Value::String(Arc::from(name.as_str())));
                    self.chunk().emit_end(line);
                } else {
                    for arg in args.iter() {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    self.emit(Op::NULL);
                }
            }
            n if n.starts_with("symbol_resolve_or_throw:") || n.starts_with("symbol_require:") => {
                // Both shapes assert the symbol resolves and raise the
                // profile-named exception when it does not. They differ only in
                // what they hand back:
                //   symbol_resolve_or_throw — the resolved symbol itself.
                //   symbol_require          — the NAME, for languages whose
                //                             surface represents a type by its
                //                             name (Java: `X.class` is a
                //                             string, so `Class.forName` must
                //                             agree with it).
                let (prefix, yields_name) = match n.starts_with("symbol_require:") {
                    true => ("symbol_require:", true),
                    false => ("symbol_resolve_or_throw:", false),
                };
                let exception_name = &n[prefix.len()..];
                if let Some(Expression {
                    kind: ExprKind::Lit(Literal::Str(name)),
                    ..
                }) = args.first()
                {
                    for arg in args.iter().skip(1) {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    let global_name = self.canon_type_global(name);
                    self.emit_constructor_global_ref(&global_name, name);
                    let line = self.line;
                    let message = name.clone();
                    crate::primitives::dynamic_symbols::emit_throw_if_unresolved(
                        self.chunk(),
                        exception_name,
                        &message,
                        line,
                    );
                    if yields_name {
                        self.emit(Op::DROP);
                        self.emit_const(Value::String(Arc::from(name.as_str())));
                    }
                } else {
                    // A computed name has no global to read at compile time;
                    // leave the argument as the value so behaviour is unchanged
                    // rather than wrongly throwing.
                    if let Some(arg) = args.first() {
                        self.compile_expr(arg)?;
                    } else {
                        self.emit(Op::NULL);
                    }
                    for arg in args.iter().skip(1) {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                }
            }
            n if n == "symbol_exists" || n.starts_with("symbol_exists:") => {
                let expected_kind = n
                    .strip_prefix("symbol_exists:")
                    .and_then(crate::primitives::reflection::ReflectKind::from_name);
                if let Some(Expression {
                    kind: ExprKind::Lit(Literal::Str(name)),
                    ..
                }) = args.first()
                {
                    let consult_resolvers = !matches!(
                        args.get(1).map(|expr| &expr.kind),
                        Some(ExprKind::Lit(Literal::Bool(false)))
                    );
                    for arg in args.iter().skip(1) {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    let global_name = self.canon_type_global(name);
                    if consult_resolvers {
                        self.emit_constructor_global_ref(&global_name, name);
                    } else {
                        let idx = self.str_const(&global_name);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                    }
                    let line = self.line;
                    crate::primitives::dynamic_symbols::emit_symbol_kind_test(
                        self.chunk(),
                        expected_kind,
                        line,
                    );
                } else {
                    // A computed name has no global to read at compile time.
                    for arg in args.iter() {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_define" => {
                if args.len() < 2 {
                    self.emit_const(Value::Bool(false));
                } else if let ExprKind::Lit(Literal::Str(name)) = &args[0].kind {
                    if let Some(ignore_case) = args.get(2) {
                        self.compile_expr(ignore_case)?;
                        self.emit(Op::DROP);
                    }
                    self.compile_expr(args[1])?;
                    let global_name = self.canon(name);
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.defined_globals.insert(global_name);
                    self.emit_const(Value::Bool(true));
                } else {
                    self.compile_expr(args[0])?;
                    self.emit(Op::DROP);
                    self.compile_expr(args[1])?;
                    self.emit(Op::DROP);
                    if let Some(ignore_case) = args.get(2) {
                        self.compile_expr(ignore_case)?;
                        self.emit(Op::DROP);
                    }
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_callable" => {
                // PHP `is_callable` matches functions and Closure
                // instances. ref_typeof on Function / HostFunction
                // returns "function" — compare via DYN_EQ.
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    fn_call!(self, "ecma:value", "typeof", 1);
                    self.emit_const(Value::String(Arc::from("function")));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_version_compare" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.emit_common("dotnet.version_parse", 1, line);
                    self.compile_expr(args[1])?;
                    self.emit_common("dotnet.version_parse", 1, line);
                    self.emit_common("dotnet.version_compare", 2, line);

                    let cmp_slot = self.define_local("__php_version_compare_cmp");
                    self.emit_u16(Op::LOCAL_SET, cmp_slot);

                    if let Some(operator) = args.get(2) {
                        let op_slot = self.define_local("__php_version_compare_op");
                        self.compile_expr(operator)?;
                        self.emit_u16(Op::LOCAL_SET, op_slot);

                        let result_slot = self.define_local("__php_version_compare_result");
                        self.emit_const(Value::Bool(false));
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        let matched_slot = self.define_local("__php_version_compare_matched");
                        self.emit_const(Value::Bool(false));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        type CmpFn = fn(&mut Chunk, u32);
                        let cmp_ops: &[(&str, CmpFn)] = &[
                            ("<", crate::primitives::ops::emit_dyn_lt as CmpFn),
                            ("lt", crate::primitives::ops::emit_dyn_lt),
                            ("<=", crate::primitives::ops::emit_dyn_le),
                            ("le", crate::primitives::ops::emit_dyn_le),
                            (">", crate::primitives::ops::emit_dyn_gt),
                            ("gt", crate::primitives::ops::emit_dyn_gt),
                            (">=", crate::primitives::ops::emit_dyn_ge),
                            ("ge", crate::primitives::ops::emit_dyn_ge),
                            ("==", crate::primitives::ops::emit_dyn_eq),
                            ("=", crate::primitives::ops::emit_dyn_eq),
                            ("eq", crate::primitives::ops::emit_dyn_eq),
                            ("!=", crate::primitives::ops::emit_dyn_ne),
                            ("<>", crate::primitives::ops::emit_dyn_ne),
                            ("ne", crate::primitives::ops::emit_dyn_ne),
                        ];
                        for (op_text, compare_fn) in cmp_ops {
                            self.emit_u16(Op::LOCAL_GET, matched_slot);
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            };
                            self.emit(Op::I32_EQZ);
                            self.chunk().emit_if(line);
                            self.emit_u16(Op::LOCAL_GET, op_slot);
                            self.emit_const(Value::String(Arc::from(*op_text)));
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                            };
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            };
                            self.chunk().emit_if(line);
                            self.emit_u16(Op::LOCAL_GET, cmp_slot);
                            self.emit_const(Value::F64(0.0));
                            {
                                let line = self.line;
                                compare_fn(self.chunk(), line);
                            };
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.emit_const(Value::Bool(true));
                            self.emit_u16(Op::LOCAL_SET, matched_slot);
                            self.chunk().emit_end(line);
                            self.chunk().emit_end(line);
                        }

                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    } else {
                        self.emit_u16(Op::LOCAL_GET, cmp_slot);
                    }
                } else {
                    self.emit_const(Value::I32(0));
                }
            }
            "php_printf" => {
                if args.is_empty() {
                    self.emit_const(Value::I32(0));
                } else {
                    let result_slot = self.define_local("__php_printf_result");
                    // PHP printf writes raw bytes to stdout — no newline.
                    // WASI 0.3 stream surface, NOT wasi:logging.log
                    // (one line record per call).
                    let write_idx = self.import("wasi:cli/stdout", "write-via-stream");
                    let rd_slot = self.define_local("__php_printf_rd");
                    let wr_slot = self.define_local("__php_printf_wr");

                    for arg in args {
                        self.compile_expr(arg)?;
                    }
                    self.emit_common("sprintf.format", args.len() as u8, line);
                    self.emit_common("php.echo_stringify", 1, line);
                    self.emit_u16(Op::LOCAL_SET, result_slot);

                    common::io::emit_write_stdout_with_imports(
                        self.chunk(),
                        write_idx,
                        rd_slot,
                        wr_slot,
                        line,
                        |c| c.emit_op_u16(Op::LOCAL_GET, result_slot, line),
                    );

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    common::strings::emit_length(self.chunk(), line);
                }
            }
            "php_vprintf" => {
                if args.len() < 2 {
                    self.emit_const(Value::I32(0));
                } else {
                    let result_slot = self.define_local("__php_vprintf_result");
                    // Raw stdout bytes via the 0.3 stream, same as php_printf.
                    let write_idx = self.import("wasi:cli/stdout", "write-via-stream");
                    let rd_slot = self.define_local("__php_vprintf_rd");
                    let wr_slot = self.define_local("__php_vprintf_wr");

                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.emit_common("sprintf.format_array", 2, line);
                    self.emit_common("php.echo_stringify", 1, line);
                    self.emit_u16(Op::LOCAL_SET, result_slot);

                    common::io::emit_write_stdout_with_imports(
                        self.chunk(),
                        write_idx,
                        rd_slot,
                        wr_slot,
                        line,
                        |c| c.emit_op_u16(Op::LOCAL_GET, result_slot, line),
                    );

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    common::strings::emit_length(self.chunk(), line);
                }
            }
            "php_vsprintf" => {
                if args.len() < 2 {
                    self.emit_const(Value::String(Arc::from("")));
                } else {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.emit_common("sprintf.format_array", 2, line);
                }
            }
            "php_register_shutdown_function" => {
                for arg in args {
                    self.compile_expr(arg)?;
                    self.emit(Op::DROP);
                }
                self.emit(Op::NULL);
            }
            // The host environment's time zone — ONE clock shared by every
            // language, so PHP's `date_default_timezone_set`, Java's
            // `TimeZone.setDefault` and .NET's `TimeZoneInfo` all write the
            // value `Intl`/`Date`/`SystemTimeZoneIdentifier` read back.
            // Language-neutral names: any profile can bind these.
            "default_timezone_set" => {
                if args.is_empty() {
                    self.emit_const(Value::Bool(false));
                } else {
                    self.compile_expr(&args[0])?;
                    for arg in &args[1..] {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    let import = self
                        .chunk()
                        .add_import("ecma:intl/timezone", "setSystemIdentifier");
                    self.chunk().emit_call(import, 1, 0);
                }
            }
            "default_timezone_get" => {
                for arg in args {
                    self.compile_expr(arg)?;
                    self.emit(Op::DROP);
                }
                let import = self
                    .chunk()
                    .add_import("ecma:intl/timezone", "systemIdentifier");
                self.chunk().emit_call(import, 0, 0);
            }
            "php_rsort" => {
                // PHP `rsort($arr)` — descending in-place sort. Compose
                // from the existing runtime helper: `sort_in_place(arr)` for the
                // ascending sort, then `array_reverse` for descending.
                // PHP arrays are JS arrays in our model, so the sort +
                // reverse mutate the same backing storage the caller's
                // variable points to.
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    let arr_slot = self.define_local("__php_rsort_arr");
                    self.emit_u16(Op::LOCAL_SET, arr_slot);
                    let helper = self.str_const("__vybe_sort_in_place");
                    self.emit_u16(Op::GLOBAL_GET, helper);
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    common::collections::emit_reverse(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                    self.emit(Op::NULL);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "right" => {
                // Right(s, n) → substring(s, len(s) - n, len(s))
                // Direct opcodes — no host call. Mirrors the `left`
                // intrinsic shape; goes through `common::strings`
                // emitters so the underlying provider (str_substring
                // opcode) stays the single source of truth.
                if args.len() >= 2 {
                    // Stash s and n in scratch slots so we can use len(s)
                    // and n twice (compute start = len - n, end = len).
                    let s_slot = self.define_local("__right_s");
                    let n_slot = self.define_local("__right_n");
                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, s_slot);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, n_slot);
                    // substring(s, len(s) - n, len(s))
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    // start = len(s) - n
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, n_slot);
                    self.emit(Op::I32_SUB);
                    let start_slot = self.define_local("__right_start");
                    self.emit_u16(Op::LOCAL_SET, start_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    inst!(self, core_wasm::i32_const, 0);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.chunk().emit_end(line);
                    // end = len(s)
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "php_substr" => {
                if args.len() >= 2 {
                    let str_slot = self.define_local("__php_substr_s");
                    let start_slot = self.define_local("__php_substr_start");
                    let len_slot = self.define_local("__php_substr_len");
                    let end_slot = self.define_local("__php_substr_end");
                    let length_slot = self.define_local("__php_substr_length");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, str_slot);

                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, start_slot);

                    self.emit_u16(Op::LOCAL_GET, str_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, len_slot);

                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, len_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_SET, start_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_u16(Op::LOCAL_SET, start_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_u16(Op::LOCAL_GET, len_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, len_slot);
                    self.emit_u16(Op::LOCAL_SET, start_slot);
                    self.chunk().emit_end(line);

                    if args.len() >= 3 {
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit_u16(Op::LOCAL_SET, length_slot);

                        self.emit_u16(Op::LOCAL_GET, length_slot);
                        inst!(self, core_wasm::i32_const, 0);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                        };
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, len_slot);
                        self.emit_u16(Op::LOCAL_GET, length_slot);
                        self.emit(Op::I32_ADD);
                        self.emit_u16(Op::LOCAL_SET, end_slot);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.emit_u16(Op::LOCAL_GET, length_slot);
                        self.emit(Op::I32_ADD);
                        self.emit_u16(Op::LOCAL_SET, end_slot);
                        self.chunk().emit_end(line);
                    } else {
                        self.emit_u16(Op::LOCAL_GET, len_slot);
                        self.emit_u16(Op::LOCAL_SET, end_slot);
                    }

                    self.emit_u16(Op::LOCAL_GET, end_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_u16(Op::LOCAL_SET, end_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, end_slot);
                    self.emit_u16(Op::LOCAL_GET, len_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, len_slot);
                    self.emit_u16(Op::LOCAL_SET, end_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, end_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_u16(Op::LOCAL_SET, end_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, str_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_u16(Op::LOCAL_GET, end_slot);

                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            // `php_strpos` / `strtr` relocated to the PHP string adapter
            // (`languages/php/emitter/string_adapter.rs`, routed via
            // `common:php.strpos` / `common:php.strtr`). PHP-specific runtime
            // semantics do not belong in the shared compiler.
            "php_str_contains" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);

                    self.compile_expr(args[1])?;
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);

                    fn_call!(self, "ecma:string", "includes", 2);
                } else {
                    inst!(self, core_wasm::bool_const, false);
                }
            }
            "php_str_starts_with" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);

                    self.compile_expr(args[1])?;
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);

                    fn_call!(self, "ecma:string", "startsWith", 2);
                } else {
                    inst!(self, core_wasm::bool_const, false);
                }
            }
            "php_str_ends_with" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);

                    self.compile_expr(args[1])?;
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);

                    fn_call!(self, "ecma:string", "endsWith", 2);
                } else {
                    inst!(self, core_wasm::bool_const, false);
                }
            }
            "php_array_search" => {
                if args.len() >= 2 {
                    let idx_slot = self.define_local("__php_array_search_idx");
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::collections::emit_index_of(&mut self.chunks, self.current, line);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);

                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.chunk().emit_else(line);
                    inst!(self, core_wasm::bool_const, false);
                    self.chunk().emit_end(line);
                } else {
                    inst!(self, core_wasm::bool_const, false);
                }
            }
            "php_array_slice" => {
                if args.len() >= 2 {
                    let arr_slot = self.define_local("__php_array_slice_arr");
                    let start_slot = self.define_local("__php_array_slice_start");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, arr_slot);

                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, start_slot);

                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);

                    if args.len() >= 3 {
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::I32_ADD);
                    } else {
                        self.emit_const(Value::I32(i32::MAX));
                    }

                    common::collections::emit_slice(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "php_range" => {
                if args.len() >= 2 {
                    let start_slot = self.define_local("__php_range_start");
                    let end_slot = self.define_local("__php_range_end");
                    let step_slot = self.define_local("__php_range_step");
                    let stop_slot = self.define_local("__php_range_stop");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, start_slot);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, end_slot);

                    if args.len() >= 3 {
                        self.compile_expr(args[2])?;
                    } else {
                        inst!(self, core_wasm::i32_const, 1);
                    }
                    self.emit_u16(Op::LOCAL_SET, step_slot);

                    self.emit_u16(Op::LOCAL_GET, step_slot);
                    self.emit_u16(Op::LOCAL_GET, end_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                    self.emit_u16(Op::LOCAL_SET, stop_slot);

                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_u16(Op::LOCAL_GET, stop_slot);
                    self.emit_u16(Op::LOCAL_GET, step_slot);
                    common::collections::emit_range(&mut self.chunks, self.current, 3, false, line);
                } else {
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                }
            }
            "php_print_expr" => {
                if let Some(arg) = args.first() {
                    let log_idx = self.import("wasi:logging/logging", "log");
                    self.compile_expr(arg)?;
                    self.emit_common("php.echo_stringify", 1, line);
                    common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);
                    inst!(self, core_wasm::i32_const, 1);
                } else {
                    inst!(self, core_wasm::i32_const, 1);
                }
            }
            "string_isnullorempty" => {
                // String.IsNullOrEmpty(s) → s is null OR str_length(s) == 0.
                // Compile s, dup, ref_is_null → if true return true, else
                // str_length == 0.
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    // [s]
                    inst!(self, core_wasm::dup);
                    // [s, s]
                    self.emit(Op::REF_IS_NULL);
                    // [s, is_null]
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit(Op::DROP);
                    inst!(self, core_wasm::bool_const, true);
                    self.chunk().emit_else(line);
                    // not null branch: [s] → str_length → cmp 0
                    common::strings::emit_length(self.chunk(), line);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    self.chunk().emit_end(line);
                } else {
                    inst!(self, core_wasm::bool_const, true);
                }
            }
            "mid" | "mid_1based" => {
                // Mid(s, start[, len]) — 1-based
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB); // start0
                    if args.len() >= 3 {
                        inst!(self, core_wasm::dup);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::I32_ADD); // start0 + length
                    } else {
                        self.emit_const(Value::I32(0x7FFF_FFFF));
                    }
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "number_isnan" => {
                self.compile_expr(args[0])?;
                inst!(self, core_wasm::dup);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                };
            }
            "number_isfinite" => {
                self.compile_expr(args[0])?;
                common::math::emit_abs(self.chunk(), line);
                self.emit_const(Value::F64(f64::MAX));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_le(self.chunk(), line);
                };
            }
            "number_isinteger" => {
                self.compile_expr(args[0])?;
                inst!(self, core_wasm::dup);
                self.emit(Op::F64_TRUNC);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
            }
            "map_size" => {
                self.compile_expr(args[0])?;
                common::dict::emit_keys(&mut self.chunks, self.current, line);
                common::collections::emit_len(&mut self.chunks, self.current, line);
            }
            "array_at" => {
                // .at() supports negative indices for both arrays and strings.
                // Receiver is already on stack from value method dispatch.
                // `Array.prototype.at` per ECMA-262 §23.1.3.1.
                if args.len() >= 1 {
                    self.compile_expr(args[0])?;
                    let idx = self.import("ecma:array", "at");
                    self.emit_host_call(idx, 2);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "instr" => {
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                } else if args.len() == 3 {
                    let start_slot = self.define_local("__instr_start");
                    self.compile_expr(args[0])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, start_slot);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_const(Value::I32(0x7FFF_FFFF));
                    common::strings::emit_substring(self.chunk(), line);
                    self.compile_expr(args[2])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    let idx_slot = self.define_local("__instr_idx");
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    inst!(self, core_wasm::i32_const, 0);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit(Op::I32_ADD);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                    self.chunk().emit_end(line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "strcomp" => {
                if args.len() >= 2 {
                    let left_slot = self.define_local("__strcomp_left");
                    let right_slot = self.define_local("__strcomp_right");
                    let text_slot = self.define_local("__strcomp_text");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    if let Some(compare_arg) = args.get(2) {
                        self.compile_expr(compare_arg)?;
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                    } else {
                        inst!(self, core_wasm::bool_const, false);
                    }
                    self.emit_u16(Op::LOCAL_SET, text_slot);

                    self.emit_u16(Op::LOCAL_GET, text_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, left_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, right_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    fn_call!(self, "wasm:js-string", "compare", 2);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, left_slot);
                    self.emit_u16(Op::LOCAL_GET, right_slot);
                    fn_call!(self, "wasm:js-string", "compare", 2);
                    self.chunk().emit_end(line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "instrrev" => {
                if args.len() >= 2 {
                    if args.len() >= 3 {
                        let source_slot = self.define_local("__instrrev_source");
                        let start_slot = self.define_local("__instrrev_start");
                        self.compile_expr(args[0])?;
                        self.emit_u16(Op::LOCAL_SET, source_slot);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit_u16(Op::LOCAL_SET, start_slot);

                        self.emit_u16(Op::LOCAL_GET, source_slot);
                        inst!(self, core_wasm::i32_const, 0);
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        common::strings::emit_substring(self.chunk(), line);
                        self.compile_expr(args[1])?;
                        common::strings::emit_last_index_of(self.chunk(), line);
                        let idx_slot = self.define_local("__instrrev_idx");
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        inst!(self, core_wasm::i32_const, 0);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                        };
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_if_value(line);
                        inst!(self, core_wasm::i32_const, 0);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::I32_ADD);
                        self.chunk().emit_end(line);
                    } else {
                        self.compile_expr(args[0])?;
                        self.compile_expr(args[1])?;
                        common::strings::emit_last_index_of(self.chunk(), line);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::I32_ADD);
                    }
                } else {
                    self.emit(Op::NULL);
                }
            }
            "fortran_index" => {
                if args.len() >= 2 {
                    let source_slot = self.define_local("__fortran_index_source");
                    let search_slot = self.define_local("__fortran_index_search");
                    let back_slot = self.define_local("__fortran_index_back");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, source_slot);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, search_slot);

                    if let Some(back_arg) = args.get(2) {
                        self.compile_expr(back_arg)?;
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                    } else {
                        inst!(self, core_wasm::bool_const, false);
                    }
                    self.emit_u16(Op::LOCAL_SET, back_slot);

                    self.emit_u16(Op::LOCAL_GET, back_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);

                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    self.emit_u16(Op::LOCAL_GET, search_slot);
                    common::strings::emit_last_index_of(self.chunk(), line);

                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    self.emit_u16(Op::LOCAL_GET, search_slot);
                    common::strings::emit_index_of(self.chunk(), line);
                    self.chunk().emit_end(line);

                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "fortran_scan" => {
                self.emit_fortran_scan_like(args, false)?;
            }
            "fortran_verify" => {
                self.emit_fortran_scan_like(args, true)?;
            }
            "replace" => {
                if args.len() >= 3 {
                    let source_slot = self.define_local("__vb_replace_source");
                    let find_slot = self.define_local("__vb_replace_find");
                    let repl_slot = self.define_local("__vb_replace_repl");
                    let start_slot = self.define_local("__vb_replace_start");
                    let count_slot = self.define_local("__vb_replace_count");
                    let text_slot = self.define_local("__vb_replace_text");
                    let result_slot = self.define_local("__vb_replace_result");
                    let remaining_slot = self.define_local("__vb_replace_remaining");
                    let find_cmp_slot = self.define_local("__vb_replace_find_cmp");
                    let current_cmp_slot = self.define_local("__vb_replace_current_cmp");
                    let find_len_slot = self.define_local("__vb_replace_find_len");
                    let idx_slot = self.define_local("__vb_replace_idx");
                    let replaced_slot = self.define_local("__vb_replace_done");
                    let prefix_slot = self.define_local("__vb_replace_prefix_end");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, source_slot);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, find_slot);

                    self.compile_expr(args[2])?;
                    self.emit_u16(Op::LOCAL_SET, repl_slot);

                    if let Some(start_arg) = args.get(3) {
                        self.compile_expr(start_arg)?;
                        common::convert::emit_to_int(self.chunk(), line);
                    } else {
                        self.emit_const(Value::I32(0));
                    }
                    self.emit_u16(Op::LOCAL_SET, start_slot);

                    if let Some(count_arg) = args.get(4) {
                        self.compile_expr(count_arg)?;
                        common::convert::emit_to_int(self.chunk(), line);
                    } else if args.get(3).is_some() {
                        self.emit_const(Value::I32(1));
                    } else {
                        self.emit_const(Value::I32(-1));
                    }
                    self.emit_u16(Op::LOCAL_SET, count_slot);

                    if let Some(compare_arg) = args.get(5) {
                        self.compile_expr(compare_arg)?;
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                    } else {
                        inst!(self, core_wasm::bool_const, false);
                    }
                    self.emit_u16(Op::LOCAL_SET, text_slot);

                    self.emit_u16(Op::LOCAL_GET, find_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, find_len_slot);

                    self.emit_u16(Op::LOCAL_GET, text_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, find_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, find_slot);
                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_SET, find_cmp_slot);

                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_u16(Op::LOCAL_SET, prefix_slot);

                    self.emit_u16(Op::LOCAL_GET, prefix_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if(line);
                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_u16(Op::LOCAL_SET, prefix_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_u16(Op::LOCAL_GET, prefix_slot);
                    common::strings::emit_substring(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, result_slot);

                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    if args.get(3).is_some() && args.get(4).is_none() {
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::I32_SUB);
                    } else {
                        self.emit_u16(Op::LOCAL_GET, prefix_slot);
                    }
                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, remaining_slot);

                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_u16(Op::LOCAL_SET, replaced_slot);

                    self.emit_u16(Op::LOCAL_GET, find_len_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    self.chunk().emit_else(line);

                    let exit_block = self.chunk().emit_block(line);
                    let (loop_patch, _) = self.chunk().emit_loop_s(line);

                    self.emit_u16(Op::LOCAL_GET, count_slot);
                    self.emit_const(Value::I32(0));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, replaced_slot);
                    self.emit_u16(Op::LOCAL_GET, count_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_br_if(2, line);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, text_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_SET, current_cmp_slot);

                    self.emit_u16(Op::LOCAL_GET, current_cmp_slot);
                    self.emit_u16(Op::LOCAL_GET, find_cmp_slot);
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);

                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_br_if(1, line);

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    common::strings::emit_substring(self.chunk(), line);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                    self.emit_u16(Op::LOCAL_GET, repl_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                    self.emit_u16(Op::LOCAL_SET, result_slot);

                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, find_len_slot);
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, remaining_slot);

                    self.emit_u16(Op::LOCAL_GET, replaced_slot);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_SET, replaced_slot);

                    self.chunk().emit_br(0, line);
                    self.chunk().emit_end(line);
                    self.chunk().patch_loop(loop_patch);
                    self.chunk().emit_end(line);
                    self.chunk().patch_block(exit_block);

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                    self.chunk().emit_end(line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "filter" => {
                if args.len() >= 2 {
                    let arr_slot = self.define_local("__vb_filter_arr");
                    let match_slot = self.define_local("__vb_filter_match");
                    let include_slot = self.define_local("__vb_filter_include");
                    let text_slot = self.define_local("__vb_filter_text");
                    let result_slot = self.define_local("__vb_filter_result");
                    let idx_slot = self.define_local("__vb_filter_idx");
                    let elem_slot = self.define_local("__vb_filter_elem");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, arr_slot);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, match_slot);

                    if let Some(include_arg) = args.get(2) {
                        self.compile_expr(include_arg)?;
                    } else {
                        inst!(self, core_wasm::bool_const, true);
                    }
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit_u16(Op::LOCAL_SET, include_slot);

                    if let Some(compare_arg) = args.get(3) {
                        self.compile_expr(compare_arg)?;
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                    } else {
                        inst!(self, core_wasm::bool_const, false);
                    }
                    self.emit_u16(Op::LOCAL_SET, text_slot);

                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                    self.emit_u16(Op::LOCAL_SET, result_slot);

                    let state = common::loops::emit_for_in_start(
                        &mut self.chunks,
                        self.current,
                        arr_slot,
                        idx_slot,
                        line,
                    );
                    self.emit_u16(Op::LOCAL_SET, elem_slot);

                    self.emit_u16(Op::LOCAL_GET, text_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);

                    self.emit_u16(Op::LOCAL_GET, elem_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, match_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    common::strings::emit_index_of(self.chunk(), line);

                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, elem_slot);
                    self.emit_u16(Op::LOCAL_GET, match_slot);
                    common::strings::emit_index_of(self.chunk(), line);
                    self.chunk().emit_end(line);

                    inst!(self, core_wasm::i32_const, 0);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                    };

                    self.emit_u16(Op::LOCAL_GET, include_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };

                    let if_block = self.chunks[self.current].emit_block(line);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    };
                    self.chunks[self.current].emit_br_if(0, line);

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit_u16(Op::LOCAL_GET, elem_slot);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);

                    self.chunks[self.current].emit_end(line);
                    self.chunks[self.current].patch_block(if_block);

                    common::loops::emit_for_in_end(
                        &mut self.chunks,
                        self.current,
                        idx_slot,
                        state,
                        line,
                    );
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "split" => {
                let source_slot = self.define_local("__vb_split_source");
                let delim_slot = self.define_local("__vb_split_delim");
                let delim_cmp_slot = self.define_local("__vb_split_delim_cmp");
                let limit_slot = self.define_local("__vb_split_limit");
                let text_slot = self.define_local("__vb_split_text");
                let result_slot = self.define_local("__vb_split_result");
                let count_slot = self.define_local("__vb_split_count");
                let remaining_slot = self.define_local("__vb_split_remaining");
                let cmp_slot = self.define_local("__vb_split_cmp");
                let delim_len_slot = self.define_local("__vb_split_delim_len");
                let idx_slot = self.define_local("__vb_split_idx");

                self.compile_expr(args[0])?;
                self.emit_u16(Op::LOCAL_SET, source_slot);

                if let Some(delim_arg) = args.get(1) {
                    self.compile_expr(delim_arg)?;
                } else {
                    self.emit_const(Value::String(Arc::from(" ")));
                }
                self.emit_u16(Op::LOCAL_SET, delim_slot);

                if let Some(limit_arg) = args.get(2) {
                    self.compile_expr(limit_arg)?;
                    common::convert::emit_to_int(self.chunk(), line);
                } else {
                    self.emit_const(Value::I32(-1));
                }
                self.emit_u16(Op::LOCAL_SET, limit_slot);

                if let Some(compare_arg) = args.get(3) {
                    self.compile_expr(compare_arg)?;
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                } else {
                    inst!(self, core_wasm::bool_const, false);
                }
                self.emit_u16(Op::LOCAL_SET, text_slot);

                self.emit_u16(Op::LOCAL_GET, delim_slot);
                common::strings::emit_length(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, delim_len_slot);

                self.emit_u16(Op::LOCAL_GET, text_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if_value(line);
                self.emit_u16(Op::LOCAL_GET, delim_slot);
                common::strings::emit_to_lower(self.chunk(), line);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, delim_slot);
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_SET, delim_cmp_slot);

                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                self.emit_u16(Op::LOCAL_SET, result_slot);

                inst!(self, core_wasm::i32_const, 0);
                self.emit_u16(Op::LOCAL_SET, count_slot);

                self.emit_u16(Op::LOCAL_GET, source_slot);
                self.emit_u16(Op::LOCAL_SET, remaining_slot);

                self.emit_u16(Op::LOCAL_GET, delim_len_slot);
                inst!(self, core_wasm::i32_const, 0);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_u16(Op::LOCAL_GET, source_slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
                self.chunk().emit_else(line);

                let exit_block = self.chunk().emit_block(line);
                let (loop_patch, _) = self.chunk().emit_loop_s(line);

                self.emit_u16(Op::LOCAL_GET, limit_slot);
                self.emit_const(Value::I32(0));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_u16(Op::LOCAL_GET, count_slot);
                self.emit_u16(Op::LOCAL_GET, limit_slot);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_SUB);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_br_if(2, line);
                self.chunk().emit_end(line);

                self.emit_u16(Op::LOCAL_GET, text_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if_value(line);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                common::strings::emit_to_lower(self.chunk(), line);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_SET, cmp_slot);

                self.emit_u16(Op::LOCAL_GET, cmp_slot);
                self.emit_u16(Op::LOCAL_GET, delim_cmp_slot);
                common::strings::emit_index_of(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, idx_slot);

                self.emit_u16(Op::LOCAL_GET, idx_slot);
                inst!(self, core_wasm::i32_const, 0);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_br_if(1, line);

                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                inst!(self, core_wasm::i32_const, 0);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                common::strings::emit_substring(self.chunk(), line);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, count_slot);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_SET, count_slot);

                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, delim_len_slot);
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                common::strings::emit_length(self.chunk(), line);
                common::strings::emit_substring(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, remaining_slot);

                self.chunk().emit_br(0, line);
                self.chunk().emit_end(line);
                self.chunk().patch_loop(loop_patch);
                self.chunk().emit_end(line);
                self.chunk().patch_block(exit_block);

                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.chunk().emit_end(line);
            }
            "join" => {
                // Two callers:
                //   - Intrinsic (`Join(arr, sep)`): args = [arr, sep],
                //     no receiver pre-pushed.
                //   - Value-method (`arr.join(sep)`): receiver `arr`
                //     already on stack, args = [sep].
                // Disambiguate by argc: 2 args → intrinsic shape; 1 arg
                // → value-method shape (only sep to push).
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                } else if args.len() == 1 {
                    self.compile_expr(args[0])?;
                } else {
                    self.emit_const(Value::String(Arc::from(",")));
                }
                {
                    let l = self.line;
                    common::collections::emit_join(&mut self.chunks, self.current, l);
                }
            }

            // ── Pascal ordinal/array intrinsics (canonical compiler_common ops) ──
            "high" => {
                // High(arr) → __len__(arr) - 1
                self.compile_expr(args[0])?;
                common::collections::emit_len(&mut self.chunks, self.current, line);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_SUB);
            }
            "low" => {
                // Low(arr) → 0 (always 0 for dynamic arrays in our VM)
                inst!(self, core_wasm::i32_const, 0);
            }
            "succ" => {
                // Succ(x) → x + 1
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(1.0));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                };
            }
            "pred" => {
                // Pred(x) → x - 1
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(1.0));
                self.emit(Op::F64_SUB);
            }
            "sqr" => {
                // Sqr(x) → x * x (square, NOT square root)
                self.compile_expr(args[0])?;
                inst!(self, core_wasm::dup);
                self.emit(Op::F64_MUL);
            }
            "assigned" => {
                // Assigned(x) → x is not null
                self.compile_expr(args[0])?;
                self.emit(Op::NULL);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                };
            }
            "sizeof" => {
                // SizeOf(x) → 4 (boxed value)
                self.compile_expr(args[0])?;
                self.emit(Op::DROP);
                self.emit_const(Value::I32(4));
            }
            "classname" => {
                // ClassName(obj) → obj.__type
                self.compile_expr(args[0])?;
                let idx = self.str_const("__type");
                self.emit_u16(Op::STRUCT_GET, idx);
            }
            "pos" => {
                // Pos(substr, s) → IndexOf(s, substr) + 1 (Pascal 1-based)
                if args.len() == 2 {
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "copy" => {
                // Copy(s, start, len) → substring(s, start-1, start-1+len) — Pascal 1-based
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB);
                    if args.len() >= 3 {
                        inst!(self, core_wasm::dup);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::I32_ADD);
                    } else {
                        self.emit_const(Value::I32(0x7FFF_FFFF));
                    }
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "leftstr" => {
                // LeftStr(s, n) → substring(s, 0, n)
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    inst!(self, core_wasm::i32_const, 0);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "str_concat" => {
                // Concat(a, b, c, ...) → a + b + c + ... using compiler_common::strings
                if args.is_empty() {
                    self.emit_const(Value::String(Arc::from("")));
                } else {
                    self.compile_expr(args[0])?;
                    for a in &args[1..] {
                        self.compile_expr(a)?;
                        common::strings::emit_str_concat(self.chunk(), line);
                    }
                }
            }
            "rightstr" => {
                // RightStr(s, n) → substring(s, len(s)-n, len(s))
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    let s_slot = self.define_local("__rs_s");
                    self.emit_u16(Op::LOCAL_SET, s_slot);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit(Op::I32_SUB);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }

            // ── String compositions of ecma:string primitives ──────────
            //
            // Each of these compiles inline so ecma:string.padStart,
            // ecma:string.toUpperCase, etc. are the single source of
            // truth for semantics. The compositions are well-known JS
            // idioms — see comments per arm.
            "zfill" => {
                // Python str.zfill(width) → padStart(width, "0").
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::String(Arc::from("0")));
                    let idx = self.import("ecma:string", "padStart");
                    self.emit_host_call(idx, 3);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "partition" | "rpartition" => {
                // `s.partition(sep)` → `(before, sep, after)`, split on the first
                // (rpartition: last) occurrence; `(s, '', '')` when not found.
                if args.len() >= 2 {
                    let from_right = name == "rpartition";
                    let s = self.define_local("__pt_s");
                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, s);
                    let sep = self.define_local("__pt_sep");
                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, sep);
                    // i = (r)indexOf(s, sep) as i32
                    self.emit_u16(Op::LOCAL_GET, s);
                    self.emit_u16(Op::LOCAL_GET, sep);
                    if from_right {
                        common::strings::emit_last_index_of(self.chunk(), line);
                    } else {
                        common::strings::emit_index_of(self.chunk(), line);
                    }
                    let to_f64 = self.import("wasm:js-number", "toF64");
                    self.emit_host_call(to_f64, 1);
                    self.emit(Op::I32_TRUNC_SAT_F64_S);
                    let i = self.define_local("__pt_i");
                    self.emit_u16(Op::LOCAL_SET, i);
                    self.emit_u16(Op::LOCAL_GET, i);
                    self.emit_const(Value::I32(0));
                    self.emit(Op::I32_LT_S);
                    self.chunk().emit_if_value(line);
                    // Not found: partition → (s, '', ''); rpartition → ('', '', s)
                    // (the unmatched whole string sits on the far side of the
                    // search direction). Universal partition semantics.
                    if from_right {
                        self.emit_const(Value::String(Arc::from("")));
                        self.emit_const(Value::String(Arc::from("")));
                        self.emit_u16(Op::LOCAL_GET, s);
                    } else {
                        self.emit_u16(Op::LOCAL_GET, s);
                        self.emit_const(Value::String(Arc::from("")));
                        self.emit_const(Value::String(Arc::from("")));
                    }
                    common::tuples::emit_tuple(&mut self.chunks, self.current, 3, line);
                    self.chunk().emit_else(line);
                    // before = s[0:i]
                    self.emit_u16(Op::LOCAL_GET, s);
                    self.emit_const(Value::I32(0));
                    self.emit_u16(Op::LOCAL_GET, i);
                    common::strings::emit_substring(self.chunk(), line);
                    // sep
                    self.emit_u16(Op::LOCAL_GET, sep);
                    // after = s[i+len(sep):]
                    self.emit_u16(Op::LOCAL_GET, s);
                    self.emit_u16(Op::LOCAL_GET, i);
                    self.emit_u16(Op::LOCAL_GET, sep);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit(Op::I32_ADD);
                    self.emit_const(Value::I32(0x7FFF_FFFF));
                    common::strings::emit_substring(self.chunk(), line);
                    common::tuples::emit_tuple(&mut self.chunks, self.current, 3, line);
                    self.chunk().emit_end(line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "capitalize" => {
                // Python/Ruby `s.capitalize()` → s[0].toUpperCase() +
                // s.slice(1).toLowerCase(). Compose via ecma:string.
                if let Some(arg) = args.first() {
                    let s_slot = self.define_local("__cap_s");
                    self.compile_expr(arg)?;
                    self.emit_u16(Op::LOCAL_SET, s_slot);
                    // first char upper
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_const(Value::I32(1));
                    common::strings::emit_substring(self.chunk(), line);
                    let upper_idx = self.import("ecma:string", "toUpperCase");
                    self.emit_host_call(upper_idx, 1);
                    // rest lower
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    self.emit_const(Value::I32(1));
                    self.emit_const(Value::I32(0x7FFF_FFFF));
                    common::strings::emit_substring(self.chunk(), line);
                    let lower_idx = self.import("ecma:string", "toLowerCase");
                    self.emit_host_call(lower_idx, 1);
                    // concat
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                } else {
                    self.emit(Op::NULL);
                }
            }
            "center" => {
                // Python str.center(width, fill?) — pad symmetrically.
                // Compose: padStart(ceil((w + len)/2), fill).padEnd(w, fill).
                if args.len() >= 2 {
                    let s_slot = self.define_local("__cen_s");
                    let w_slot = self.define_local("__cen_w");
                    let pad_slot = self.define_local("__cen_pad");
                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, s_slot);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, w_slot);
                    if args.len() >= 3 {
                        self.compile_expr(args[2])?;
                    } else {
                        self.emit_const(Value::String(Arc::from(" ")));
                    }
                    self.emit_u16(Op::LOCAL_SET, pad_slot);
                    // Step 1: padStart with target = (w + len) / 2 + len_remainder
                    // For simplicity: padStart with (w + len + 1)/2.
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    // target = (w + len + 1) / 2
                    self.emit_u16(Op::LOCAL_GET, w_slot);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit(Op::I32_ADD);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                    self.emit_const(Value::I32(2));
                    self.emit(Op::I32_DIV_S);
                    self.emit_u16(Op::LOCAL_GET, pad_slot);
                    let pad_start = self.import("ecma:string", "padStart");
                    self.emit_host_call(pad_start, 3);
                    // Step 2: padEnd to full width.
                    self.emit_u16(Op::LOCAL_GET, w_slot);
                    self.emit_u16(Op::LOCAL_GET, pad_slot);
                    let pad_end = self.import("ecma:string", "padEnd");
                    self.emit_host_call(pad_end, 3);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "count" => {
                // Python `s.count(sub)` / PHP `substr_count($s, $sub)` —
                // count non-overlapping occurrences. Compose:
                // s.split(sub).length - 1.
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    let split_idx = self.import("ecma:string", "split");
                    self.emit_host_call(split_idx, 2);
                    // Length of the host array from split → i32 (the polymorphic
                    // emit_len uses GC array.len, which misreads a host array).
                    let arr_len = self.import("ecma:array", "length");
                    self.emit_host_call(arr_len, 1);
                    let to_f64 = self.import("wasm:js-number", "toF64");
                    self.emit_host_call(to_f64, 1);
                    self.emit(Op::I32_TRUNC_SAT_F64_S);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "chop" => {
                // Ruby `s.chop` — drop last char. Compose: s.slice(0, len(s)-1).
                if let Some(arg) = args.first() {
                    let s_slot = self.define_local("__chop_s");
                    self.compile_expr(arg)?;
                    self.emit_u16(Op::LOCAL_SET, s_slot);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "chars" => {
                // Ruby/PHP `s.chars` — array of single-char strings.
                // Compose: s.split("").
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    self.emit_const(Value::String(Arc::from("")));
                    let split_idx = self.import("ecma:string", "split");
                    self.emit_host_call(split_idx, 2);
                } else {
                    let empty_arr = self.import("ecma:array", "new");
                    self.emit_host_call(empty_arr, 0);
                }
            }

            // ── Numeric conversion intrinsics ─────────────────────────
            //
            // VB / Pascal / Python `cint` / `int(x)` / `clng` — coerce
            // to a number and round to nearest-even so midpoint cases
            // line up with VB's Round semantics.
            "cint" | "clng" => {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let value_slot = self.define_local("__cint_value");
                    let result_slot = self.define_local("__cint_result");
                    let handled_slot = self.define_local("__cint_handled");
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit_const(Value::Bool(false));
                    self.emit_u16(Op::LOCAL_SET, handled_slot);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    fn_call!(self, "ecma:value", "typeof", 1);
                    self.emit_const(Value::String(Arc::from("string")));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    inst!(self, core_wasm::i32_const, 0);
                    fn_call!(self, "wasm:js-string", "charCodeAt", 2);
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.emit_const(Value::Bool(true));
                    self.emit_u16(Op::LOCAL_SET, handled_slot);
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, handled_slot);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let num = self.import("ecma:number", "Number");
                    self.emit_host_call(num, 1);
                    let rounded_value_slot = self.define_local("__cint_rounded_value");
                    let floor_slot = self.define_local("__cint_floor");
                    let ceil_slot = self.define_local("__cint_ceil");
                    let frac_slot = self.define_local("__cint_frac");

                    self.emit_u16(Op::LOCAL_SET, rounded_value_slot);
                    if self.profile.name == "vb" {
                        self.emit_u16(Op::LOCAL_GET, rounded_value_slot);
                        self.emit_u16(Op::LOCAL_GET, rounded_value_slot);
                        self.emit(Op::F64_NE);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::String(Arc::from(
                            "Conversion from string to type 'Integer' is not valid.",
                        )));
                        self.emit_js_exception_ctor_from_message_value("InvalidCastException")?;
                        common::errors::emit_throw(self.chunk(), line);
                        self.chunk().emit_end(line);
                    }

                    self.emit_u16(Op::LOCAL_GET, rounded_value_slot);
                    self.emit(Op::F64_FLOOR);
                    self.emit_u16(Op::LOCAL_SET, floor_slot);

                    self.emit_u16(Op::LOCAL_GET, rounded_value_slot);
                    self.emit(Op::F64_CEIL);
                    self.emit_u16(Op::LOCAL_SET, ceil_slot);

                    self.emit_u16(Op::LOCAL_GET, rounded_value_slot);
                    self.emit_u16(Op::LOCAL_GET, floor_slot);
                    self.emit(Op::F64_SUB);
                    self.emit_u16(Op::LOCAL_SET, frac_slot);

                    self.emit_u16(Op::LOCAL_GET, frac_slot);
                    self.emit_const(Value::F64(0.5));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, floor_slot);

                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, frac_slot);
                    self.emit_const(Value::F64(0.5));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, ceil_slot);

                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, floor_slot);
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_AND);
                    self.emit_const(Value::I32(0));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, floor_slot);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, ceil_slot);
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);

                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                } else {
                    self.emit_const(Value::F64(0.0));
                }
            }

            // VB `hex(n)` / `Hex$` — uppercase hex string.
            // ECMA composition: `Number(n).toString(16).toUpperCase()`.
            // `Number.prototype.toString` is called via a method
            // dispatch on the numeric receiver; `String.prototype.
            // toUpperCase` likewise.
            "hex" => {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let num = self.import("ecma:number", "Number");
                    self.emit_host_call(num, 1);
                    // Number(n).toString(16)
                    self.emit_const(Value::F64(16.0));
                    let to_str = self.import("ecma:number", "toString");
                    self.emit_host_call(to_str, 2);
                    // .toUpperCase()
                    let upper = self.import("ecma:string", "toUpperCase");
                    self.emit_host_call(upper, 1);
                } else {
                    self.emit_const(Value::String(Arc::from("0")));
                }
            }

            // VB `oct(n)` / `Oct$` — octal string.
            "oct" => {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let num = self.import("ecma:number", "Number");
                    self.emit_host_call(num, 1);
                    self.emit_const(Value::F64(8.0));
                    let to_str = self.import("ecma:number", "toString");
                    self.emit_host_call(to_str, 2);
                } else {
                    self.emit_const(Value::String(Arc::from("0")));
                }
            }

            _ => {
                self.emit(Op::NULL);
            }
        }
        Ok(())
    }
}
