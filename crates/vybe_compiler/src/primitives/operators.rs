//! Operator, arithmetic-coercion, and compound-assignment compilation.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — same pattern as
//! `builtins.rs`/`calls.rs`. Methods are `pub(super)` so the core compile
//! paths in `mod.rs` and sibling files can reach them.

use super::*;

impl Compiler {
    pub(super) fn emit_to_primitive(&mut self, hint: &str) {
        let value_slot = self.define_local("__to_primitive_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        let idx = self.import("ecma:value", "toPrimitive");
        self.emit_const(Value::String(Arc::from(hint)));
        self.emit_host_call(idx, 2);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.chunk().emit_end(line);
    }

    /// JS profile: coerce both top-of-stack operands via the
    /// ToPrimitive polyfill, then to_f64 via the VM's existing
    /// `Value::as_f64` once the operand is no longer an Object.
    /// Used for `-`, `*`, `/`. Passes hint="number" per ECMA §7.1.4
    /// step 1 (ToNumber unboxes Objects with hint=number first).
    pub(super) fn coerce_top_two_to_number(&mut self) {
        let t_b = self.define_local("__binop_b");
        self.emit_u16(Op::LOCAL_SET, t_b);
        // a on top → coerce
        self.emit_to_primitive("number");
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit_to_primitive("number");
    }

    /// Runtime-polymorphic numeric binary op for dynamically-typed
    /// languages. Stack `[a, b]` → `[result]`. At runtime, if BOTH operands
    /// are BigInt, calls `ecma:bigint.<bigint_fn>` (which returns a
    /// `Value::BigInt`, so the result stays BigInt-typed through a chain);
    /// otherwise runs the SAME number path the static route would
    /// (`coerce_top_two_to_number` — honouring `valueOf`/ToPrimitive — then
    /// the f64 op), so non-BigInt behaviour is byte-for-byte unchanged.
    pub(super) fn emit_js_dynamic_arith(&mut self, bigint_fn: &str, number_op: NumberArith) {
        let b_slot = self.define_local("__dynarith_b");
        self.emit_u16(Op::LOCAL_SET, b_slot);
        let a_slot = self.define_local("__dynarith_a");
        self.emit_u16(Op::LOCAL_SET, a_slot);

        let test_bi = self.import("wasm:js-bigint", "test");
        self.emit_u16(Op::LOCAL_GET, a_slot);
        self.emit_host_call(test_bi, 1);
        self.emit_u16(Op::LOCAL_GET, b_slot);
        self.emit_host_call(test_bi, 1);
        self.emit(Op::I32_AND);
        let line = self.line;
        self.chunk().emit_if_value(line);

        // both BigInt → ecma:bigint.<fn>(a, b)
        self.emit_u16(Op::LOCAL_GET, a_slot);
        self.emit_u16(Op::LOCAL_GET, b_slot);
        let bi = self.import("ecma:bigint", bigint_fn);
        self.emit_host_call(bi, 2);

        self.chunk().emit_else(line);
        // number path — identical to the static route.
        self.emit_u16(Op::LOCAL_GET, a_slot);
        self.emit_u16(Op::LOCAL_GET, b_slot);
        self.coerce_top_two_to_number();
        match number_op {
            NumberArith::Sub => self.emit(Op::F64_SUB),
            NumberArith::Mul => self.emit(Op::F64_MUL),
            NumberArith::Div => self.emit(Op::F64_DIV),
            NumberArith::Mod => {
                let l = self.line;
                common::math::emit_c_fmod(self.chunk(), l);
            }
        }
        let line = self.line;
        self.chunk().emit_end(line);
    }

    /// §13.15.2 strict-mode assignment: a false [[Set]] result throws
    /// TypeError. Consumes the Bool left by the strict proxy set dispatch.
    pub(super) fn emit_strict_set_failure_check(&mut self) -> Result<(), String> {
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);
        self.emit_const(Value::String(Arc::from(
            "'set' on proxy: trap returned falsish",
        )));
        self.emit_js_exception_ctor_from_message_value("TypeError")?;
        common::errors::emit_throw(self.chunk(), line);
        self.chunk().emit_end(line);
        Ok(())
    }

    /// `++`/`--` step. ECMA §13.4: ToNumeric keeps the operand's numeric
    /// type — a BigInt operand steps by 1n (result stays BigInt), anything
    /// else by Number 1. Profiles without BigInt keep the plain number
    /// path byte-for-byte.
    pub(crate) fn emit_step_by_one(&mut self, add: bool) {
        if self.profile.has_ecma_bigint {
            let slot = self.define_local("__step_v");
            self.emit_u16(Op::LOCAL_SET, slot);
            let test_bi = self.import("wasm:js-bigint", "test");
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_host_call(test_bi, 1);
            let line = self.line;
            self.chunk().emit_if_value(line);
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_const(Value::F64(1.0));
            let bi = self.import("ecma:bigint", if add { "add" } else { "sub" });
            self.emit_host_call(bi, 2);
            self.chunk().emit_else(line);
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_const(Value::F64(1.0));
            if add {
                crate::primitives::ops::emit_dyn_add(self.chunk(), line);
            } else {
                self.emit(Op::F64_SUB);
            }
            self.chunk().emit_end(line);
        } else {
            self.emit_const(Value::F64(1.0));
            if add {
                let line = self.line;
                crate::primitives::ops::emit_dyn_add(self.chunk(), line);
            } else {
                self.emit(Op::F64_SUB);
            }
        }
    }

    /// JS profile: ToPrimitive(hint=number) on both operands. Used
    /// before DYN_LT / DYN_GT / DYN_LE / DYN_GE so string-string lex
    /// compare and Date/valueOf-overriding instances both work.
    pub(super) fn coerce_top_two_to_primitive(&mut self) {
        let t_b = self.define_local("__cmpop_b");
        self.emit_u16(Op::LOCAL_SET, t_b);
        self.emit_to_primitive("number");
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit_to_primitive("number");
    }

    pub(super) fn emit_pascal_relational_compare(&mut self, cmp_fn: fn(&mut Chunk, u32)) {
        let t_b = self.define_local("__pas_cmp_b");
        let t_a = self.define_local("__pas_cmp_a");
        self.emit_u16(Op::LOCAL_SET, t_b);
        self.emit_u16(Op::LOCAL_SET, t_a);

        self.emit_u16(Op::LOCAL_GET, t_a);
        fn_call!(self, "ecma:value", "typeof", 1);
        self.emit_const(Value::String(Arc::from("string")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, t_b);
        fn_call!(self, "ecma:value", "typeof", 1);
        self.emit_const(Value::String(Arc::from("string")));
        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit_u16(Op::LOCAL_GET, t_b);
        fn_call!(self, "wasm:js-string", "compare", 2);
        self.emit_const(Value::I32(0));
        cmp_fn(self.chunk(), line);

        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit_u16(Op::LOCAL_GET, t_b);
        cmp_fn(self.chunk(), line);
        self.chunk().emit_end(line);

        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit_u16(Op::LOCAL_GET, t_b);
        cmp_fn(self.chunk(), line);
        self.chunk().emit_end(line);
    }

    /// JS profile: ToPrimitive(hint=default) on both operands. Used
    /// before DYN_ADD per ECMA §13.15.4 — the `+` operator picks the
    /// "default" hint, which gives valueOf the first shot and falls
    /// back to toString.
    pub(super) fn coerce_top_two_to_default_primitive(&mut self) {
        let t_b = self.define_local("__addop_b");
        self.emit_u16(Op::LOCAL_SET, t_b);
        self.emit_to_primitive("default");
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit_to_primitive("default");
    }

    pub(super) fn emit_js_add_string_concat_from_locals(&mut self, lhs_slot: u16, rhs_slot: u16) {
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_to_primitive("string");
        self.emit_const(Value::String(Arc::from("")));
        common::strings::emit_str_concat(self.chunk(), line);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit_to_primitive("string");
        self.emit_const(Value::String(Arc::from("")));
        common::strings::emit_str_concat(self.chunk(), line);
        fn_call!(self, "wasm:js-string", "concat", 2);
    }

    pub(super) fn emit_js_add_numeric_from_locals(&mut self, lhs_slot: u16, rhs_slot: u16) {
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_const(Value::F64(f64::NAN));
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let rhs_line = self.line;
        self.chunk().emit_if_value(rhs_line);
        self.emit_const(Value::F64(f64::NAN));
        self.chunk().emit_else(rhs_line);
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };
        self.chunk().emit_end(rhs_line);
        self.chunk().emit_end(line);
    }

    pub(super) fn emit_python_contains_from_locals(
        &mut self,
        container_slot: u16,
        needle_slot: u16,
    ) {
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, container_slot);
        fn_call!(self, "wasm:js-string", "test", 1);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, container_slot);
        self.emit_u16(Op::LOCAL_GET, needle_slot);
        fn_call!(self, "ecma:string", "includes", 2);
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, container_slot);
        let is_array = self.import("ecma:array", "isArray");
        self.chunk().emit_call(is_array, 1, line);
        inst!(self, core_wasm::i32_const, 0);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
        };
        let array_line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), array_line);
        self.chunk().emit_if_value(array_line);
        self.emit_u16(Op::LOCAL_GET, container_slot);
        self.emit_u16(Op::LOCAL_GET, needle_slot);
        common::collections::emit_contains(&mut self.chunks, self.current, line);
        self.chunk().emit_else(array_line);
        self.emit_u16(Op::LOCAL_GET, container_slot);
        self.emit_u16(Op::LOCAL_GET, needle_slot);
        let has_in = self.import("ecma:object", "hasIn");
        self.chunk().emit_call(has_in, 2, line);
        self.chunk().emit_end(array_line);
        self.chunk().emit_end(line);
    }

    #[allow(dead_code)]
    pub(super) fn emit_js_add(&mut self) {
        let rhs_slot = self.define_local("__js_add_rhs");
        let lhs_slot = self.define_local("__js_add_lhs");
        self.emit_u16(Op::LOCAL_SET, rhs_slot);
        self.emit_u16(Op::LOCAL_SET, lhs_slot);

        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        self.emit_const(Value::String(Arc::from("string")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);
        self.emit_js_add_string_concat_from_locals(lhs_slot, rhs_slot);
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        self.emit_const(Value::String(Arc::from("string")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        };
        let rhs_line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), rhs_line);
        self.chunk().emit_if_value(rhs_line);
        self.emit_js_add_string_concat_from_locals(lhs_slot, rhs_slot);
        self.chunk().emit_else(rhs_line);
        self.emit_js_add_numeric_from_locals(lhs_slot, rhs_slot);
        self.chunk().emit_end(rhs_line);
        self.chunk().emit_end(line);
    }

    /// Try a user operator method on the left operand, else `fallback`.
    /// Stack: `[lhs, rhs]` → `[result]`.
    ///
    /// Operator overloading is a method call the operands opt into, so the
    /// dispatch lives here next to every other operator lowering rather
    /// than in any one language's adapter — a class that defines
    /// `operator +` (Dart), `__add__` (Python) or `op_Addition` (C#)
    /// normalises to the same bound method name, so this reaches all of
    /// them, including across languages.
    /// Try the user operator bound to `slot` on the operands, else `fallback`.
    ///
    /// Takes a SLOT, not a spelling, so an operator cannot be dispatched by
    /// name here even by accident — and the same code reaches Python's
    /// `__add__`, Dart's `operator +` and C#'s `operator+` because all three
    /// fill `ProtocolSlot::Add`.
    fn emit_rich_binop(&mut self, slot: vybe_ast::ProtocolSlot, fallback: fn(&mut Chunk, u32)) {
        let line = self.line;
        let rhs_slot = self.define_local("__rich_op_rhs");
        let lhs_slot = self.define_local("__rich_op_lhs");
        self.emit_u16(Op::LOCAL_SET, rhs_slot);
        self.emit_u16(Op::LOCAL_SET, lhs_slot);
        common::expressions::emit_rich_arithmetic(
            self.chunk(),
            lhs_slot,
            rhs_slot,
            &vybe_ast::protocol_slot_key(slot),
            fallback,
            line,
        );
    }

    /// Try a user unary-operator method on the operand, else `fallback`.
    /// Stack: `[operand]` → `[result]`.
    pub(super) fn emit_rich_unary(
        &mut self,
        protocol_slot: vybe_ast::ProtocolSlot,
        fallback: fn(&mut Chunk, u32),
    ) {
        let line = self.line;
        let slot = self.define_local("__rich_op_operand");
        self.emit_u16(Op::LOCAL_SET, slot);
        common::expressions::emit_rich_unary(
            self.chunk(),
            slot,
            &vybe_ast::protocol_slot_key(protocol_slot),
            fallback,
            line,
        );
    }

    pub(super) fn compile_binop(&mut self, op: &BinOp) {
        match op {
            BinOp::Add => {
                // `dynamic_add`: JS-style `+` — concatenates when either
                // operand is a string, otherwise adds numerically. PHP,
                // Python, Lua, etc. use `.` / `..` / other operators for
                // string concat, so `+` is purely numeric and coerces
                // string operands (`"2026" + 4 == 2030`). `F64_ADD`
                // coerces both sides via `Value::as_f64()`; `DYN_ADD`
                // has the JS-style string-concat special case.
                if self.profile.dynamic_add && self.profile.name != "cobol" {
                    // JS profile: ECMA §13.15.4 — call ToPrimitive on
                    // both operands with hint "default" before adding.
                    // The polyfill returns the operand unchanged for
                    // primitives (fast path) and unboxes Objects via
                    // their valueOf/toString chain (Date, custom
                    // valueOf, class instances).
                    if self.profile.ecma_operator_coercion {
                        let idx = self.import("ecma:value", "add");
                        self.emit_host_call(idx, 2);
                        return;
                    }
                    // A user `operator +` / `__add__` defines `+` for its own
                    // type; `emit_dyn_add` would coerce the operand instead
                    // and never consult it. Falls through to the same
                    // dynamic add for every non-object operand.
                    if self.uses_rich_operators() {
                        self.emit_rich_binop(
                            vybe_ast::ProtocolSlot::Add,
                            crate::primitives::ops::emit_dyn_add,
                        );
                        return;
                    }
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                } else if self.is_php_profile() {
                    // PHP `+` on arrays = union (first-wins merge).
                    // Check at runtime if both are objects → union.
                    let a_slot = self.define_local("__php_add_a");
                    let b_slot = self.define_local("__php_add_b");
                    self.emit_u16(Op::LOCAL_SET, b_slot);
                    self.emit_u16(Op::LOCAL_SET, a_slot);
                    self.emit_u16(Op::LOCAL_GET, a_slot);
                    inst!(self, recipes::is_object);
                    self.emit_u16(Op::LOCAL_GET, b_slot);
                    inst!(self, recipes::is_object);
                    self.emit(Op::I32_AND);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    // Both objects → array union: iterate b's entries,
                    // set on a copy of a only if key doesn't exist
                    self.emit_u16(Op::LOCAL_GET, a_slot);
                    self.emit_u16(Op::LOCAL_GET, b_slot);
                    self.emit_common("php.array_union", 2, line);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, a_slot);
                    self.emit_u16(Op::LOCAL_GET, b_slot);
                    self.emit(Op::F64_ADD);
                    self.chunk().emit_end(line);
                } else {
                    self.emit(Op::F64_ADD);
                }
            }
            BinOp::Sub => {
                if self.profile.dynamic_numeric_dispatch {
                    self.emit_js_dynamic_arith("sub", NumberArith::Sub);
                } else {
                    if self.profile.ecma_operator_coercion {
                        self.coerce_top_two_to_number();
                    }
                    self.emit(Op::F64_SUB);
                }
            }
            BinOp::Mul => {
                if self.profile.dynamic_numeric_dispatch {
                    self.emit_js_dynamic_arith("mul", NumberArith::Mul);
                } else {
                    if self.profile.ecma_operator_coercion {
                        self.coerce_top_two_to_number();
                    }
                    self.emit(Op::F64_MUL);
                }
            }
            BinOp::Div => {
                if self.profile.dynamic_numeric_dispatch {
                    self.emit_js_dynamic_arith("div", NumberArith::Div);
                } else {
                    if self.profile.ecma_operator_coercion {
                        self.coerce_top_two_to_number();
                    }
                    self.emit(Op::F64_DIV);
                }
            }
            BinOp::IDiv => {
                self.emit(Op::F64_DIV);
                let l = self.line;
                common::math::emit_trunc(self.chunk(), l);
            }
            BinOp::FloorDiv => {
                self.emit(Op::F64_DIV);
                let l = self.line;
                common::math::emit_floor(self.chunk(), l);
            }
            BinOp::Mod => {
                let l = self.line;
                if self.is_python_profile() {
                    common::math::emit_python_floor_mod(self.chunk(), l);
                } else if self.profile.dynamic_numeric_dispatch {
                    self.emit_js_dynamic_arith("rem", NumberArith::Mod);
                } else {
                    common::math::emit_c_fmod(self.chunk(), l);
                }
            }
            BinOp::Pow => {
                let l = self.line;
                common::math::emit_pow(self.chunk(), l);
            }
            BinOp::Eq => {
                if self.profile.abstract_equality {
                    let idx = self.import("ecma:value", "abstractEq");
                    self.emit_host_call(idx, 2);
                } else if self.uses_rich_comparison() {
                    // Dispatch to a user `__eq__` (or cross-language alias) with
                    // the receiver, falling back to structural equality.
                    let right_slot = self.define_local("__rich_eq_rhs");
                    let left_slot = self.define_local("__rich_eq_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    // Structural-equality fallback is the language's registered
                    // `value_eq` hook (present iff the language wants deep value
                    // equality — Python tuples/dicts, Dart records), else plain
                    // reference/primitive equality. Hook presence is the property;
                    // never a profile-name check.
                    let eq_fallback = vybe_bytecode::registry::hooks(&self.profile.name)
                        .value_eq
                        .unwrap_or(crate::primitives::ops::emit_dyn_eq);
                    common::expressions::emit_rich_compare_locals(
                        self.chunk(),
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq),
                        eq_fallback,
                        line,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                } else {
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                        if self.profile.materialize_bool_results {
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                        }
                    };
                }
            }
            BinOp::NotEq => {
                if self.profile.abstract_equality {
                    let idx = self.import("ecma:value", "abstractNe");
                    self.emit_host_call(idx, 2);
                } else if self.uses_rich_comparison() {
                    // `a != b` == not (a `__eq__` b) — dispatch `__eq__`, negate.
                    let right_slot = self.define_local("__rich_ne_rhs");
                    let left_slot = self.define_local("__rich_ne_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    // Structural-equality fallback is the language's registered
                    // `value_eq` hook (present iff the language wants deep value
                    // equality — Python tuples/dicts, Dart records), else plain
                    // reference/primitive equality. Hook presence is the property;
                    // never a profile-name check.
                    let eq_fallback = vybe_bytecode::registry::hooks(&self.profile.name)
                        .value_eq
                        .unwrap_or(crate::primitives::ops::emit_dyn_eq);
                    common::expressions::emit_rich_compare_locals(
                        self.chunk(),
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq),
                        eq_fallback,
                        line,
                    );
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                } else {
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                        if self.profile.materialize_bool_results {
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                        }
                    };
                }
            }
            BinOp::StrictEq => {
                let line = self.line;
                if self.profile.ecma_operator_coercion {
                    crate::primitives::ops::emit_js_strict_eq(self.chunk(), line);
                } else {
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                }
                if self.profile.ecma_boolean_operators {
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                } else if self.profile.materialize_bool_results {
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                }
            }
            BinOp::StrictNotEq => {
                // JS !==: negate of ===.
                {
                    let line = self.line;
                    if self.profile.ecma_operator_coercion {
                        crate::primitives::ops::emit_js_strict_eq(self.chunk(), line);
                        self.emit(Op::I32_EQZ);
                    } else {
                        crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                    }
                    if self.profile.ecma_boolean_operators {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    } else if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::Lt => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive();
                } else if self.profile.string_aware_relational {
                    let line = self.line;
                    vybe_bytecode::registry::hooks(&self.profile.name)
                        .relational_compare
                        .unwrap()(
                        self.chunk(), crate::primitives::ops::emit_dyn_lt, line
                    );
                    return;
                }
                if self.profile.name == "pascal" {
                    self.emit_pascal_relational_compare(crate::primitives::ops::emit_dyn_lt);
                } else if self.profile.ecma_operator_coercion {
                    // ECMA-262 §7.2.13 — NaN-safe ToNumber on mixed operands.
                    let line = self.line;
                    crate::primitives::ops::emit_js_lt(self.chunk(), line);
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                } else if self.is_php_profile() {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                } else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(
                        self.chunk(),
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Lt),
                        crate::primitives::ops::emit_dyn_lt,
                        line,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::Gt => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive();
                } else if self.profile.string_aware_relational {
                    let line = self.line;
                    vybe_bytecode::registry::hooks(&self.profile.name)
                        .relational_compare
                        .unwrap()(
                        self.chunk(), crate::primitives::ops::emit_dyn_gt, line
                    );
                    return;
                }
                if self.profile.name == "pascal" {
                    self.emit_pascal_relational_compare(crate::primitives::ops::emit_dyn_gt);
                } else if self.profile.ecma_operator_coercion {
                    // ECMA-262 §7.2.13 — NaN-safe ToNumber on mixed operands.
                    let line = self.line;
                    crate::primitives::ops::emit_js_gt(self.chunk(), line);
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                } else if self.is_php_profile() {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                } else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(
                        self.chunk(),
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Gt),
                        crate::primitives::ops::emit_dyn_gt,
                        line,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::LtEq => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive();
                } else if self.profile.string_aware_relational {
                    let line = self.line;
                    vybe_bytecode::registry::hooks(&self.profile.name)
                        .relational_compare
                        .unwrap()(
                        self.chunk(), crate::primitives::ops::emit_dyn_le, line
                    );
                    return;
                }
                if self.profile.name == "pascal" {
                    self.emit_pascal_relational_compare(crate::primitives::ops::emit_dyn_le);
                } else if self.profile.ecma_operator_coercion {
                    // ECMA-262 §7.2.13 — NaN-safe ToNumber on mixed operands.
                    let line = self.line;
                    crate::primitives::ops::emit_js_le(self.chunk(), line);
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                } else if self.is_php_profile() {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_le(self.chunk(), line);
                } else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(
                        self.chunk(),
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Le),
                        crate::primitives::ops::emit_dyn_le,
                        line,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::GtEq => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive();
                } else if self.profile.string_aware_relational {
                    let line = self.line;
                    vybe_bytecode::registry::hooks(&self.profile.name)
                        .relational_compare
                        .unwrap()(
                        self.chunk(), crate::primitives::ops::emit_dyn_ge, line
                    );
                    return;
                }
                if self.profile.name == "pascal" {
                    self.emit_pascal_relational_compare(crate::primitives::ops::emit_dyn_ge);
                } else if self.profile.ecma_operator_coercion {
                    // ECMA-262 §7.2.13 — NaN-safe ToNumber on mixed operands.
                    let line = self.line;
                    crate::primitives::ops::emit_js_ge(self.chunk(), line);
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                } else if self.is_php_profile() {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                } else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(
                        self.chunk(),
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Ge),
                        crate::primitives::ops::emit_dyn_ge,
                        line,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::Spaceship => {
                let right_slot = self.define_local("__spaceship_rhs");
                let left_slot = self.define_local("__spaceship_lhs");
                self.emit_u16(Op::LOCAL_SET, right_slot);
                self.emit_u16(Op::LOCAL_SET, left_slot);

                self.emit_u16(Op::LOCAL_GET, left_slot);
                self.emit_u16(Op::LOCAL_GET, right_slot);
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive();
                } else if self.profile.string_aware_relational {
                    let line = self.line;
                    vybe_bytecode::registry::hooks(&self.profile.name)
                        .relational_compare
                        .unwrap()(
                        self.chunk(), crate::primitives::ops::emit_dyn_lt, line
                    );
                } else if self.profile.name == "pascal" {
                    self.emit_pascal_relational_compare(crate::primitives::ops::emit_dyn_lt);
                } else {
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                }
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                self.emit_const(Value::I32(-1));
                self.chunk().emit_else(line);

                self.emit_u16(Op::LOCAL_GET, left_slot);
                self.emit_u16(Op::LOCAL_GET, right_slot);
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive();
                } else if self.profile.string_aware_relational {
                    let line = self.line;
                    vybe_bytecode::registry::hooks(&self.profile.name)
                        .relational_compare
                        .unwrap()(
                        self.chunk(), crate::primitives::ops::emit_dyn_gt, line
                    );
                } else if self.profile.name == "pascal" {
                    self.emit_pascal_relational_compare(crate::primitives::ops::emit_dyn_gt);
                } else {
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                    };
                }
                let gt_line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), gt_line);
                self.chunk().emit_if_value(gt_line);
                self.emit_const(Value::I32(1));
                self.chunk().emit_else(gt_line);
                self.emit_const(Value::I32(0));
                self.chunk().emit_end(gt_line);
                self.chunk().emit_end(line);
            }
            BinOp::And | BinOp::Or => unreachable!(), // handled with short-circuit
            BinOp::Xor => {
                self.emit(Op::I32_XOR);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
            }
            BinOp::Eqv => {
                self.emit(Op::I32_XOR);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            }
            BinOp::Imp => {
                let rhs_slot = self.define_local("__imp_rhs");
                let lhs_slot = self.define_local("__imp_lhs");
                self.emit_u16(Op::LOCAL_SET, rhs_slot);
                self.emit_u16(Op::LOCAL_SET, lhs_slot);
                self.emit_u16(Op::LOCAL_GET, lhs_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
                self.emit_u16(Op::LOCAL_GET, rhs_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.emit(Op::I32_OR);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
            }
            BinOp::BitAnd => self.emit(Op::I32_AND),
            BinOp::BitOr => self.emit(Op::I32_OR),
            BinOp::BitXor => self.emit(Op::I32_XOR),
            BinOp::Shl => self.emit(Op::I32_SHL),
            BinOp::Shr => self.emit(Op::I32_SHR_S),
            BinOp::UShr => {
                self.emit(Op::I32_SHR_U);
                if self.profile.ecma_operator_coercion {
                    // ECMA-262 §13.10.2: `>>>` produces an unsigned 32-bit
                    // integer (Number). I32_SHR_U leaves the bit pattern
                    // in i32, but if the high bit is set (e.g. `-1 >>> 0`)
                    // the i32 → Number coercion would render as negative.
                    // Reinterpret as u32 by adding 2^32 when the i32 is
                    // negative — keeps within the f64 53-bit mantissa.
                    self.emit(Op::F64_FROM_I32);
                    let unsigned_slot = self.define_local("__js_ushr_result");
                    self.emit_u16(Op::LOCAL_SET, unsigned_slot);
                    self.emit_u16(Op::LOCAL_GET, unsigned_slot);
                    self.emit_const(Value::F64(0.0));
                    self.emit(Op::F64_LT);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, unsigned_slot);
                    self.emit_const(Value::F64(4_294_967_296.0));
                    self.emit(Op::F64_ADD);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, unsigned_slot);
                    self.chunk().emit_end(line);
                }
            }
            BinOp::Concat => {
                let l = self.line;
                common::strings::emit_str_concat_coercing(self.chunk(), l);
            }
            BinOp::In => {
                if self.is_python_profile() {
                    let t_y = self.define_local("__py_in_y");
                    let t_x = self.define_local("__py_in_x");
                    self.emit_u16(Op::LOCAL_SET, t_y);
                    self.emit_u16(Op::LOCAL_SET, t_x);
                    self.emit_python_contains_from_locals(t_y, t_x);
                    return;
                }

                if self.profile.name == "pascal" {
                    let t_set = self.define_local("__pascal_in_set");
                    let t_value = self.define_local("__pascal_in_value");
                    self.emit_u16(Op::LOCAL_SET, t_set);
                    self.emit_u16(Op::LOCAL_SET, t_value);
                    let helper = self.str_const("__vybe_pascal_set_contains");
                    self.emit_u16(Op::GLOBAL_GET, helper);
                    self.emit_u16(Op::LOCAL_GET, t_value);
                    self.emit_u16(Op::LOCAL_GET, t_set);
                    self.emit_u8(Op::CALL_REF, 2);
                    return;
                }

                // `x in y` — JS: is `x` a property KEY of `y` (not a value).
                // ECMA-262 §13.10.1 walks the prototype chain. PHP
                // `in_array` / Python `key in dict` are own-only and
                // route through their language profiles separately.
                //
                // Walker stack: `[x, y]`. hasIn expects `[y, x]`.
                let l = self.line;
                let t_y = self.define_local("__in_y");
                let t_x = self.define_local("__in_x");
                self.emit_u16(Op::LOCAL_SET, t_y);
                self.emit_u16(Op::LOCAL_SET, t_x);

                // Proxy has-trap dispatch on the JS profile when the
                // module references `Proxy`. Stack: [obj, key].
                if self.uses_proxy {
                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    vybe_bytecode::registry::hooks(&self.profile.name)
                        .proxy_has
                        .unwrap()(&mut self.chunks, self.current, l);
                    return;
                }

                self.emit_u16(Op::LOCAL_GET, t_y);
                self.emit_u16(Op::LOCAL_GET, t_x);
                // JS uses prototype-walking `hasIn`; other languages
                // (case-insensitive profiles or non-JS) keep own-only
                // `hasOwn` semantics for their `in`-shaped operators.
                let import = if self.profile.ecma_in_operator {
                    "hasIn"
                } else {
                    "hasOwn"
                };
                let idx = self.import("ecma:object", import);
                self.chunk().emit_call(idx, 2, l);
                // hasIn/hasOwn return Value::Bool — already correct for ECMA display.
            }
            BinOp::NotIn => {
                if self.is_python_profile() {
                    let t_y = self.define_local("__py_nin_y");
                    let t_x = self.define_local("__py_nin_x");
                    self.emit_u16(Op::LOCAL_SET, t_y);
                    self.emit_u16(Op::LOCAL_SET, t_x);
                    self.emit_python_contains_from_locals(t_y, t_x);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    };
                    return;
                }

                if self.profile.name == "pascal" {
                    let t_set = self.define_local("__pascal_nin_set");
                    let t_value = self.define_local("__pascal_nin_value");
                    self.emit_u16(Op::LOCAL_SET, t_set);
                    self.emit_u16(Op::LOCAL_SET, t_value);
                    let helper = self.str_const("__vybe_pascal_set_contains");
                    self.emit_u16(Op::GLOBAL_GET, helper);
                    self.emit_u16(Op::LOCAL_GET, t_value);
                    self.emit_u16(Op::LOCAL_GET, t_set);
                    self.emit_u8(Op::CALL_REF, 2);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    };
                    return;
                }

                let l = self.line;
                let t_y = self.define_local("__nin_y");
                let t_x = self.define_local("__nin_x");
                self.emit_u16(Op::LOCAL_SET, t_y);
                self.emit_u16(Op::LOCAL_SET, t_x);
                self.emit_u16(Op::LOCAL_GET, t_y);
                self.emit_u16(Op::LOCAL_GET, t_x);
                // Same key-check as `in` above — route through hasOwn.
                let idx = self.import("ecma:object", "hasOwn");
                self.chunk().emit_call(idx, 2, l);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            }
            BinOp::InstanceOf => {
                if self.class_prototype_dispatch() {
                    let rhs_slot = self.define_local("__js_instanceof_rhs");
                    let lhs_slot = self.define_local("__js_instanceof_lhs");
                    self.emit_u16(Op::LOCAL_SET, rhs_slot);
                    self.emit_u16(Op::LOCAL_SET, lhs_slot);
                    // ECMA-262 §13.10.2: `a instanceof B` first checks for
                    // `B[Symbol.hasInstance]` (canonical name `hasinstance`)
                    // and calls it as `B[hasinstance](a)` if present.
                    // Compiler-side dispatch keeps the JS method-call
                    // protocol intact (`__js_this` bound to B) — host
                    // `ctx.invoke` can't do that, so we emit the
                    // method-call inline instead of going through the
                    // host fn for this case.
                    let has_inst_key = self.str_const("hasinstance");
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.emit_u16(Op::STRUCT_GET, has_inst_key);
                    let method_slot = self.define_local("__has_inst_method");
                    self.emit_u16(Op::LOCAL_SET, method_slot);
                    self.emit_u16(Op::LOCAL_GET, method_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    let helper = self.import("ecma:value", "instanceOf");
                    self.emit_u16(Op::LOCAL_GET, lhs_slot);
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.emit_host_call(helper, 2);
                    self.chunk().emit_else(line);
                    let saved_this = self.save_js_this("__js_prev_this_hasinst");
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, method_slot);
                    self.emit_u16(Op::LOCAL_GET, lhs_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    {
                        let line = self.line;
                        // Convert dynamic result to Bool (consistent with
                        // instanceOf host fn which also returns Bool).
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    };
                    let result_slot = self.define_local("__has_inst_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.restore_js_this(saved_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.chunk().emit_end(line);
                } else {
                    // Dynamic-RHS fallback: the static `a instanceof TypeName`
                    // form is intercepted upstream in `expressions.rs` and
                    // emitted as `Op::REF_TEST` directly. This branch only
                    // fires for the rare `a instanceof <expression>` shape.
                    //
                    // Stack on entry: [val, ctor]. We string-compare
                    // `val.__type` against `ctor.name` — the same compile-time
                    // type-stamp the constructors install via `set_type_id`.
                    let l = self.line;
                    // Stack [val, ctor]. Extract the constructor's NAME string,
                    // then membership-check it against val's `__types` ancestry
                    // (falling back to the single `__type`). Inheritance-aware
                    // and shared: `is_a?`/`isinstance` normalize to this node.
                    let t_ctor = self.define_local("__io_ctor");
                    self.emit_u16(Op::LOCAL_SET, t_ctor); // [val]
                    self.emit_u16(Op::LOCAL_GET, t_ctor);
                    let name_key = self.str_const("name");
                    self.chunk().emit_op_u16(Op::STRUCT_GET, name_key, l); // [val, ctor_name]
                    common::reflection::emit_instanceof(&mut self.chunks, self.current, l);
                }
            }
            BinOp::NullCoalesce => unreachable!(), // handled in compile_expr
            BinOp::MatMul => {
                let i = self.import("ecma:math", "matmul");
                self.emit_host_call(i, 2);
            }
            BinOp::Like => {
                // The VB walker always rewrites `a Like b` to `Regex.IsMatch(pattern, a)`
                // before reaching this point, so this arm is dead for VB.
                // Stack at call site: [string, pattern]; ecma:regexp.test expects
                // (pattern, string) — callers via the walker path never reach here.
                let idx = self.import("ecma:regexp", "test");
                self.emit_host_call(idx, 2);
            }
            BinOp::Is => {
                // Reference equality
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
            }
            BinOp::IsNot => {
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Compound assignment operator emission
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_compound_op(&mut self, op: &CompoundOp) {
        match op {
            CompoundOp::Add => {
                if self.profile.dynamic_add && self.profile.name != "cobol" {
                    if self.profile.ecma_operator_coercion {
                        self.coerce_top_two_to_default_primitive();
                    }
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                } else {
                    self.emit(Op::F64_ADD);
                }
            }
            CompoundOp::Sub => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_SUB);
            }
            CompoundOp::Mul => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_MUL);
            }
            CompoundOp::Div => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_DIV);
            }
            CompoundOp::IDiv => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_DIV);
                let l = self.line;
                common::math::emit_trunc(self.chunk(), l);
            }
            CompoundOp::Mod => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                let l = self.line;
                common::math::emit_c_fmod(self.chunk(), l);
            }
            CompoundOp::Pow => {
                let l = self.line;
                common::math::emit_pow(self.chunk(), l);
            }
            CompoundOp::Concat => {
                let l = self.line;
                common::strings::emit_str_concat(self.chunk(), l);
            }
            CompoundOp::BitAnd => self.emit(Op::I32_AND),
            CompoundOp::BitOr => self.emit(Op::I32_OR),
            CompoundOp::BitXor => self.emit(Op::I32_XOR),
            CompoundOp::Shl => self.emit(Op::I32_SHL),
            CompoundOp::Shr => self.emit(Op::I32_SHR_S),
            CompoundOp::UShr => self.emit(Op::I32_SHR_U),
            CompoundOp::And => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            } // simplified
            CompoundOp::Or => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            } // simplified
            CompoundOp::NullCoalesce => {
                // a ??= b → if a is null, a = b
                // At this point both are on stack already — no-op, the whole compound assign handles it
            }
        }
    }

    pub(super) fn is_csharp_delegate_handler_expr(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lambda { .. } | ExprKind::AddressOf(_) => true,
            ExprKind::Ident(name) => {
                if self
                    .lookup_var_type_hint(name)
                    .is_some_and(Self::is_callable_type_hint)
                {
                    return true;
                }
                if self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some())
                {
                    return false;
                }
                let cname = self.canon(name);
                self.defined_functions.contains(&cname)
                    || self.defined_class_methods.contains(&cname)
            }
            ExprKind::Member { field, .. } => {
                let cname = self.canon(field);
                self.defined_functions.contains(&cname)
                    || self.defined_class_methods.contains(&cname)
            }
            ExprKind::New { args, .. } if args.len() == 1 => {
                self.is_csharp_delegate_handler_expr(&args[0].value)
            }
            _ => false,
        }
    }

    pub(super) fn assign_target_matches_expr(
        &self,
        target: &Expression,
        expr: &Expression,
    ) -> bool {
        match (&target.kind, &expr.kind) {
            (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
            (
                ExprKind::Member {
                    object: to,
                    field: tf,
                    ..
                },
                ExprKind::Member {
                    object: eo,
                    field: ef,
                    ..
                },
            ) => {
                if !tf.eq_ignore_ascii_case(ef) {
                    return false;
                }
                self.assign_target_matches_expr(to, eo)
            }
            (ExprKind::This, ExprKind::This) => true,
            _ => false,
        }
    }

    // ════════════════════════════════════════════════════════════════════════
}
