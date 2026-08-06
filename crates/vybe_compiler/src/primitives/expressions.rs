//! `compile_expr` — the `ExprKind` dispatch. The single largest
//! method in the compiler, split out so edits to other concerns
//! (calls, classes, statements) don't churn a multi-thousand-line
//! file.

use super::*;

impl Compiler {
    /// Construct a tree `Type` node generically — the common-resolver
    /// construction path (namespaceplan.md), retiring per-platform surfaces.
    ///
    /// Reorders named args by the spec's param names (the shared named-arg
    /// machinery), allocates the object, stamps `__type` + the `__types`
    /// ancestry array (so `is`/`instanceof` matches every ancestor), and stores
    /// each constructor argument into its field — an omitted optional stores an
    /// explicit `null`. Any language/platform registering a tree `Ctor`
    /// (`flutter.*`, and eventually the dotnet BCL) constructs through this ONE
    /// path. Leaves the constructed object on the stack.
    pub(crate) fn emit_tree_ctor_construction(
        &mut self,
        spec: &crate::primitives::namespaces::CtorSpec,
        args: &[crate::ast::Argument],
    ) -> Result<(), String> {
        use crate::ast::{Expression, Param, PassBy};

        let params: Vec<Param> = spec
            .params
            .iter()
            .map(|name| Param {
                name: name.clone(),
                type_hint: None,
                default: Some(Expression::null()),
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: true,
                is_nullable: true })
            .collect();
        let sig = super::CallSignature::from_params(&params);
        let ordered = self.reorder_named_args_with_signatures(args, &[sig]);

        use crate::primitives::namespaces::FieldGui;

        // A GUI control IS its element — `CtorSpec`'s own contract says "the
        // object IS the control". It used to build a plain config object and
        // defer to a runtime realizer that would create the backing control
        // from `__controlfn`/`__ops`; no such realizer exists (nothing reads
        // either key), so the control was never created and every property
        // write landed on the document instead of the element.
        //
        // Creating the element here makes the stamps below land ON it, so the
        // object carries both its class identity and its node.
        let this_slot = self.define_local("__tree_ctor_this");
        let control_type = spec.ancestry.first().cloned().unwrap_or_default();
        let is_control = spec.control_fn.is_some();
        if is_control {
            let line = self.line;
            self.emit_control_element(&control_type, 0, line);
        } else {
            self.emit_struct_new(0, 0);
        }
        self.emit_u16(Op::LOCAL_SET, this_slot);

        // Compile each constructor arg ONCE into a slot (child widgets must not
        // be built twice — once for the field, once for the op).
        let mut arg_slots = Vec::with_capacity(spec.fields.len());
        for i in 0..spec.fields.len() {
            let slot = self.define_local("__tree_ctor_arg");
            match ordered.get(i) {
                Some(arg) => self.compile_expr(&arg.value)?,
                None => self.emit_null() }
            self.emit_u16(Op::LOCAL_SET, slot);
            arg_slots.push(slot);
        }

        // __type = ancestry[0]
        if let Some(name) = spec.ancestry.first() {
            self.emit_u16(Op::LOCAL_GET, this_slot);
            self.emit_const(Value::String(std::sync::Arc::from(name.as_str())));
            let k = self.str_const("__type");
            self.emit_struct_field_op(Op::STRUCT_SET, 0, k);
        }
        // __types = full ancestry array (js_instanceof membership check)
        self.emit_u16(Op::LOCAL_GET, this_slot);
        for name in &spec.ancestry {
            self.emit_const(Value::String(std::sync::Arc::from(name.as_str())));
        }
        self.emit_array_new_fixed(0, spec.ancestry.len() as u16);
        let tk = self.str_const("__types");
        self.emit_struct_field_op(Op::STRUCT_SET, 0, tk);

        // __controlfn — the `vybe:gui` factory for this widget (`new_Label`…),
        // or null for a plain tree type. Marks a GUI-adapter widget.
        self.emit_u16(Op::LOCAL_GET, this_slot);
        match &spec.control_fn {
            Some(cf) => self.emit_const(Value::String(std::sync::Arc::from(cf.as_str()))),
            None => self.emit_null() }
        let cfk = self.str_const("__controlfn");
        self.emit_struct_field_op(Op::STRUCT_SET, 0, cfk);

        // __value_eq — mark immutable value types (Flutter ValueKey/Color/…)
        // so the language `==` compares them structurally (by __type + fields)
        // rather than by reference identity.
        if spec.value_equality {
            self.emit_u16(Op::LOCAL_GET, this_slot);
            self.emit_const(Value::Bool(true));
            let vk = self.str_const("__value_eq");
            self.emit_struct_field_op(Op::STRUCT_SET, 0, vk);
        }

        // Store each arg as a readable field (`Scaffold(appBar:x).appBar`).
        for (i, field) in spec.fields.iter().enumerate() {
            self.emit_u16(Op::LOCAL_GET, this_slot);
            self.emit_u16(Op::LOCAL_GET, arg_slots[i]);
            let fk = self.str_const(field);
            self.emit_struct_field_op(Op::STRUCT_SET, 0, fk);
        }

        // For GUI widgets, stamp __ops = [[kind,key,value],…] — the realizer's
        // instruction list. kind: 0 NestOrProp(key), 1 Children, 2 Event(name),
        // 3 Caption.
        if spec.control_fn.is_some() {
            self.emit_u16(Op::LOCAL_GET, this_slot);
            for (i, field) in spec.fields.iter().enumerate() {
                let (kind, key): (i32, &str) = match spec.field_gui.get(i) {
                    Some(FieldGui::Children) => (1, ""),
                    Some(FieldGui::Event(name)) => (2, name.as_str()),
                    Some(FieldGui::Caption) => (3, ""),
                    Some(FieldGui::NestOrProp(k)) => (0, k.as_str()),
                    None => (0, field.as_str()) };
                self.emit_const(Value::I32(kind));
                self.emit_const(Value::String(std::sync::Arc::from(key)));
                self.emit_u16(Op::LOCAL_GET, arg_slots[i]);
                self.emit_array_new_fixed(0, 3);
            }
            self.emit_array_new_fixed(0, spec.fields.len() as u16);
            let ok = self.str_const("__ops");
            self.emit_struct_field_op(Op::STRUCT_SET, 0, ok);
        }

        self.emit_u16(Op::LOCAL_GET, this_slot);
        Ok(())
    }

    pub(super) fn emit_member_get_from_value(&mut self, obj_slot: u16, field_name: &str) {
        // ECMA-262 §7.3.2 GetV(V, P): for an OBJECT receiver, `[[Get]]`
        // directly (Reflect.get, §28.1.6) — walks the prototype chain, invokes
        // a `__get_<field>` accessor with the receiver, else returns the raw
        // data property. For a PRIMITIVE receiver the spec says ToObject(V)
        // first, then `[[Get]]` on that wrapper — which means the member is
        // looked up on the primitive's INTRINSIC prototype.
        //
        // This branch used to yield `undefined` for every primitive, so
        // `Number.prototype.doubled = …; (5).doubled` read undefined while
        // `[1,2].second()` worked (arrays are ordinary objects). That is an
        // ECMA conformance bug in JS AND the cause of every Dart extension on a
        // built-in type (`extension on int` / `on String` / `on List<int>`) —
        // one gap, two languages, so the fix is here and not a per-language
        // call-site rewrite.
        //
        // `ecma:object.getPrototypeOf` already answers for primitives (it falls
        // through to `js_prototype_of` on a non-object), so ToObject needs no
        // wrapper allocation and no new host function: read the intrinsic
        // prototype, then `[[Get]]` on it. `Reflect.get` throws on a
        // non-object, so the primitive can never be passed to it directly.
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        self.chunk().emit_if_value(line);
        let idx = self.import("ecma:reflect", "get");
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(field_name)));
        self.emit_host_call(idx, 2);
        self.chunk().emit_else(line);
        self.emit_primitive_prototype_member_get(obj_slot, field_name);
        self.chunk().emit_end(line);

        self.emit_getattr_slot_probe(obj_slot, field_name);
    }

    /// The `GetAttr` role — the attribute-miss interceptor (Python
    /// `__getattr__`, PHP `__get`, JS Proxy get). Resolved by SLOT, so the
    /// spelling each language used is irrelevant here; both frontends already
    /// bind it to `ProtocolSlot::GetAttr`.
    ///
    /// ADDED after the normal read, never substituted for it: this site also
    /// serves plain maps, host objects and primitives, which carry no slot, and
    /// flexclassplan records that substituting at such a site is what broke
    /// `Len` and `Iterator`. So the read happens first and only an `undefined`
    /// result consults the slot.
    ///
    /// Whole thing is skipped unless some class in the program binds the role,
    /// so programs without it emit exactly what they did before.
    ///
    /// Stack: `[value] -> [value]`.
    fn emit_getattr_slot_probe(&mut self, obj_slot: u16, field_name: &str) {
        if !self.program_has_getattr {
            return;
        }
        // The slot's implementation is published under a key derived from the
        // slot NUMBER, so this never mentions `__getattr__` or `__get`.
        let slot_key = vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::GetAttr);
        let line = self.line;
        let value_slot = self.define_local("__getattr_value");
        let handler_slot = self.define_local("__getattr_handler");

        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        {
            let undef = self.chunk().add_import("wasm:js-undefined", "test");
            self.chunk().emit_call(undef, 1, line);
        }
        self.chunk().emit_if(line);

        // Only an object can carry the slot; `Reflect.get` throws otherwise.
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        inst!(self, recipes::is_object);
        self.chunk().emit_if(line);
        let get = self.import("ecma:reflect", "get");
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(slot_key.as_str())));
        self.emit_host_call(get, 2);
        self.emit_u16(Op::LOCAL_SET, handler_slot);

        self.emit_u16(Op::LOCAL_GET, handler_slot);
        {
            let undef = self.chunk().add_import("wasm:js-undefined", "test");
            self.chunk().emit_call(undef, 1, line);
        }
        self.chunk().emit_op(Op::I32_EQZ, line);
        self.chunk().emit_if(line);
        // handler(receiver, name)
        self.emit_u16(Op::LOCAL_GET, handler_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(field_name)));
        self.chunk().emit_op_u8(Op::CALL_REF, 2, line);
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.chunk().emit_end(line);

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, value_slot);
    }

    /// §7.3.2 GetV on a PRIMITIVE receiver: look `field_name` up on the
    /// primitive's intrinsic prototype. `null`/`undefined` have no prototype —
    /// `getPrototypeOf` yields undefined for them and `[[Get]]` on undefined
    /// would throw, so those short-circuit to `undefined` rather than raising,
    /// preserving the lenient read this path has always had.
    pub(super) fn emit_primitive_prototype_member_get(&mut self, obj_slot: u16, field_name: &str) {
        let line = self.line;
        let proto_slot = self.define_local("__js_primitive_proto");

        let get_proto = self.import("ecma:object", "getPrototypeOf");
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_host_call(get_proto, 1);
        self.emit_u16(Op::LOCAL_SET, proto_slot);

        self.emit_u16(Op::LOCAL_GET, proto_slot);
        inst!(self, recipes::is_object);
        self.chunk().emit_if_value(line);
        let reflect_get = self.import("ecma:reflect", "get");
        self.emit_u16(Op::LOCAL_GET, proto_slot);
        self.emit_const(Value::String(Arc::from(field_name)));
        self.emit_host_call(reflect_get, 2);
        self.chunk().emit_else(line);
        inst!(self, core_wasm::undefined);
        self.chunk().emit_end(line);
    }

    fn emit_js_import_meta_object(&mut self) {
        let global_name = "__js_import_meta";
        let meta_slot = self.define_local("__js_import_meta_value");

        self.emit_global_read(global_name);
        self.emit_u16(Op::LOCAL_SET, meta_slot);

        self.emit_u16(Op::LOCAL_GET, meta_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);
        common::dict::emit_new(&mut self.chunks, self.current, line);
        let init_slot = self.define_local("__js_import_meta_init");
        self.emit_u16(Op::LOCAL_SET, init_slot);

        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_const(Value::String(Arc::from("")));
        let url_key = self.str_const("url");
        self.emit_struct_field_op(Op::STRUCT_SET, 0, url_key);

        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_global_write(global_name);
        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_u16(Op::LOCAL_SET, meta_slot);

        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, meta_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);
        common::dict::emit_new(&mut self.chunks, self.current, line);
        let init_slot = self.define_local("__js_import_meta_init");
        self.emit_u16(Op::LOCAL_SET, init_slot);

        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_const(Value::String(Arc::from("")));
        let url_key = self.str_const("url");
        self.emit_struct_field_op(Op::STRUCT_SET, 0, url_key);

        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_global_write(global_name);
        self.emit_u16(Op::LOCAL_GET, init_slot);
        self.emit_u16(Op::LOCAL_SET, meta_slot);

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, meta_slot);
    }

    fn try_compile_fortran_array_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if !self.profile.array_arithmetic_elementwise {
            return Ok(false);
        }
        if !matches!(
            op,
            BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Pow
        ) {
            return Ok(false);
        }

        let left_is_array = self.expr_is_array_like(left);
        let right_is_array = self.expr_is_array_like(right);
        if !left_is_array && !right_is_array {
            return Ok(false);
        }

        let line = self.line;
        let left_slot = self.define_local("__fortran_array_binop_left");
        self.compile_expr(left)?;
        self.emit_u16(Op::LOCAL_SET, left_slot);

        let right_slot = self.define_local("__fortran_array_binop_right");
        self.compile_expr(right)?;
        self.emit_u16(Op::LOCAL_SET, right_slot);

        let iter_slot = if left_is_array { left_slot } else { right_slot };
        let result_slot = self.define_local("__fortran_array_binop_result");
        let idx_slot = self.define_local("__fortran_array_binop_idx");

        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, result_slot);

        let state = common::loops::emit_for_in_start(
            &mut self.chunks,
            self.current,
            iter_slot,
            idx_slot,
            line,
        );
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, result_slot);
        if left_is_array {
            self.emit_u16(Op::LOCAL_GET, left_slot);
            self.emit_u16(Op::LOCAL_GET, idx_slot);
            common::collections::emit_get(&mut self.chunks, self.current, line);
        } else {
            self.emit_u16(Op::LOCAL_GET, left_slot);
        }
        if right_is_array {
            self.emit_u16(Op::LOCAL_GET, right_slot);
            self.emit_u16(Op::LOCAL_GET, idx_slot);
            common::collections::emit_get(&mut self.chunks, self.current, line);
        } else {
            self.emit_u16(Op::LOCAL_GET, right_slot);
        }
        self.compile_binop(op);
        common::collections::emit_push(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, state, line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(true)
    }

    // ════════════════════════════════════════════════════════════════════════
    // Expression compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn emit_generator_entry_control(&mut self, control_slot: u16) -> Result<(), String> {
        // Detection logic lives in emitter/generators.rs; return handling
        // stays here because it needs Compiler state (finally blocks).
        self.emit_u16(Op::LOCAL_GET, control_slot);
        {
            let l = self.line;
            crate::primitives::instructions::recipes::is_object(self.chunk(), l);
        }
        {
            let l = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), l);
        }
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, control_slot);
        let marker_key = self.str_const("__vybe_generator_control");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, marker_key);
        {
            let l = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), l);
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, control_slot);
        let op_key = self.str_const("op");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, op_key);
        self.emit_const(Value::String(Arc::from("throw")));
        {
            let l = self.line;
            crate::primitives::instructions::host::emit(
                self.chunk(),
                "wasm:js-string",
                "equals",
                2,
                l,
            );
        }
        {
            let l = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), l);
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, control_slot);
        let value_key = self.str_const("value");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, value_key);
        {
            let line = self.line;
            crate::primitives::errors::emit_throw(self.chunk(), line);
        }

        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, control_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, op_key);
        self.emit_const(Value::String(Arc::from("return")));
        {
            let l = self.line;
            crate::primitives::instructions::host::emit(
                self.chunk(),
                "wasm:js-string",
                "equals",
                2,
                l,
            );
        }
        {
            let l = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), l);
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, control_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, value_key);
        self.emit_return_through_finally(1)?;

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        Ok(())
    }

    pub(super) fn emit_generator_resume_value(&mut self) -> Result<(), String> {
        let resume_slot = self.define_local("__yield_resume");
        self.emit_u16(Op::LOCAL_SET, resume_slot);

        let result_slot = self.define_local("__yield_resume_value");
        self.emit_u16(Op::LOCAL_GET, resume_slot);
        self.emit_u16(Op::LOCAL_SET, result_slot);

        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, resume_slot);
        {
            let l = self.line;
            crate::primitives::instructions::recipes::is_object(self.chunk(), l);
        }
        {
            let l = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), l);
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, resume_slot);
        let marker_key = self.str_const("__vybe_generator_control");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, marker_key);
        {
            let l = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), l);
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, resume_slot);
        let op_key = self.str_const("op");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, op_key);
        self.emit_const(Value::String(Arc::from("throw")));
        {
            let l = self.line;
            crate::primitives::instructions::host::emit(
                self.chunk(),
                "wasm:js-string",
                "equals",
                2,
                l,
            );
        }
        {
            let l = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), l);
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, resume_slot);
        let value_key = self.str_const("value");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, value_key);
        {
            let line = self.line;
            crate::primitives::errors::emit_throw(self.chunk(), line);
        }

        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, resume_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, op_key);
        self.emit_const(Value::String(Arc::from("return")));
        {
            let l = self.line;
            crate::primitives::instructions::host::emit(
                self.chunk(),
                "wasm:js-string",
                "equals",
                2,
                l,
            );
        }
        {
            let l = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), l);
        }
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, resume_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, value_key);
        self.emit_return_through_finally(1)?;

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(())
    }

    fn generator_keyed_yield_parts<'a>(
        &self,
        expr: &'a Expression,
    ) -> Option<(&'a Expression, &'a Expression)> {
        let ExprKind::Object(props) = &expr.kind else {
            return None;
        };

        let mut has_marker = false;
        let mut key_expr = None;
        let mut value_expr = None;
        for prop in props {
            if let ObjectProperty::KeyValue { key, value } = prop {
                match &key.kind {
                    ExprKind::Lit(Literal::Str(name)) if name == "__vybe_generator_yield" => {
                        has_marker = matches!(value.kind, ExprKind::Lit(Literal::Bool(true)));
                    }
                    ExprKind::Lit(Literal::Str(name)) if name == "key" => {
                        key_expr = Some(value);
                    }
                    ExprKind::Lit(Literal::Str(name)) if name == "value" => {
                        value_expr = Some(value);
                    }
                    _ => {}
                }
            }
        }

        if has_marker {
            key_expr.zip(value_expr)
        } else {
            None
        }
    }

    fn emit_generator_payload_store(&mut self) {
        self.emit_global_read("__vybe_generator_payloads");
        let store_slot = self.define_local("__gen_payload_store");
        self.emit_u16(Op::LOCAL_SET, store_slot);

        self.emit_u16(Op::LOCAL_GET, store_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);
        common::dict::emit_new(&mut self.chunks, self.current, line);
        inst!(self, core_wasm::dup);
        self.emit_u16(Op::LOCAL_SET, store_slot);
        self.emit_global_write("__vybe_generator_payloads");

        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, store_slot);
    }

    fn emit_next_generator_payload_id(&mut self) {
        self.emit_global_read("__vybe_generator_payload_next_id");
        let id_slot = self.define_local("__gen_payload_id_current");
        self.emit_u16(Op::LOCAL_SET, id_slot);

        self.emit_u16(Op::LOCAL_GET, id_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_const(Value::F64(0.0));
        self.emit_u16(Op::LOCAL_SET, id_slot);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, id_slot);
        self.emit_const(Value::F64(1.0));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_global_write("__vybe_generator_payload_next_id");

        self.emit_u16(Op::LOCAL_GET, id_slot);
    }

    pub(crate) fn emit_generator_yield_value(&mut self, yielded_slot: u16) {
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let marker_key = self.str_const("__vybe_generator_yield");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, marker_key);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let payload_id_key = self.str_const("payload_id");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, payload_id_key);
        let payload_id_slot = self.define_local("__yield_payload_id");
        self.emit_u16(Op::LOCAL_SET, payload_id_slot);

        self.emit_u16(Op::LOCAL_GET, payload_id_slot);
        self.emit(Op::REF_IS_NULL);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let value_key = self.str_const("value");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, value_key);
        self.chunk().emit_else(line);

        self.emit_generator_payload_store();
        self.emit_u16(Op::LOCAL_GET, payload_id_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);

        self.chunk().emit_end(line);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        self.chunk().emit_end(line);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        self.chunk().emit_end(line);
    }

    pub(crate) fn emit_generator_yield_key_or_fallback(
        &mut self,
        yielded_slot: u16,
        fallback_slot: Option<u16>,
    ) {
        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let marker_key = self.str_const("__vybe_generator_yield");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, marker_key);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, yielded_slot);
        let key_key = self.str_const("key");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, key_key);

        self.chunk().emit_else(line);
        if let Some(slot) = fallback_slot {
            self.emit_u16(Op::LOCAL_GET, slot);
        } else {
            self.emit_null();
        }
        self.chunk().emit_end(line);
        self.chunk().emit_else(line);
        if let Some(slot) = fallback_slot {
            self.emit_u16(Op::LOCAL_GET, slot);
        } else {
            self.emit_null();
        }
        self.chunk().emit_end(line);
    }

    pub(crate) fn compile_expr(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            // ── Literals ────────────────────────────────────────────────
            ExprKind::Lit(lit) => match lit {
                Literal::Int(n) => self.emit_const(Value::F64(*n as f64)),
                Literal::Float(n) => self.emit_const(Value::F64(*n)),
                Literal::BigInt(n) => self.emit_const(Value::bigint_i64(*n)),
                Literal::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                Literal::Char(c) => {
                    self.emit_const(Value::String(Arc::from(c.to_string().as_str())))
                }
                Literal::Bool(b) => {
                    if *b {
                        inst!(self, core_wasm::bool_const, true)
                    } else {
                        inst!(self, core_wasm::bool_const, false)
                    }
                }
                Literal::Null => self.emit_null(),
                Literal::Undefined => {
                    let l = self.line;
                    common::expressions::emit_undefined(self.chunk(), l);
                }
                Literal::Ellipsis => self.emit_null(),
                // A byte string becomes a real `Uint8Array` — the shape Python
                // `bytes` already used via `__py_bytes_new__`, now reached from
                // the AST so the value carries a STATIC type and
                // `(Bytes, slot)` bindings can resolve for it.
                Literal::Bytes(bytes) => {
                    for byte in bytes {
                        self.emit_const(Value::F64(f64::from(*byte)));
                    }
                    let line = self.line;
                    self.chunk()
                        .emit_array_new_fixed(0, bytes.len() as u16, line);
                    let idx = self.import("ecma:uint8array", "new");
                    self.emit_host_call(idx, 1);
                }
            },

            // ── Identifier ──────────────────────────────────────────────
            ExprKind::Ident(name) => {
                // JS global constants that aren't variables
                match name.as_str() {
                    "NaN" => {
                        self.emit_const(Value::F64(f64::NAN));
                        return Ok(());
                    }
                    "Infinity" => {
                        self.emit_const(Value::F64(f64::INFINITY));
                        return Ok(());
                    }
                    "__js_import_meta" if self.profile.supports_dynamic_import => {
                        self.emit_js_import_meta_object();
                        return Ok(());
                    }
                    "undefined" if self.case_sensitive => {
                        let l = self.line;
                        common::expressions::emit_undefined(self.chunk(), l);
                        return Ok(());
                    }
                    _ => {}
                }
                // `__debug__` / `__name__` moved to the profile's
                // `[namespace_constants]` — a bool and a string constant, which
                // is exactly what that table declares.
                if let (Some(function_name), Some(result_slot)) =
                    (self.current_func_name.as_ref(), self.current_result_slot)
                {
                    let matches = if self.case_sensitive {
                        name == function_name
                    } else {
                        name.eq_ignore_ascii_case(function_name)
                    };
                    if matches {
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        return Ok(());
                    }
                }
                // Local variable / parameter takes priority over implicit self field
                let is_local = self.scope().resolve(name).is_some()
                    || self.has_static_local_binding(name);

                // The owner of `Items` has to be an `ObservableCollection`, and
                // that is the namespace tree's answer, already scoped by
                // `type_scopes` — no other tree declares the type, so no
                // language gate can change who reaches here.
                if !is_local
                    && name.eq_ignore_ascii_case("Items")
                    && self
                        .current_class
                        .as_deref()
                        .and_then(|class_name| {
                            self.namespace_tree_instance_method_owner(class_name, "Items", 0)
                        })
                        .is_some_and(|owner| {
                            owner.to_ascii_lowercase().contains("observablecollection")
                        })
                {
                    if self.emit_self_ref() {
                        self.emit_common("dotnet.observable_collection_items", 1, self.line);
                        return Ok(());
                    }
                }

                // Implicit self field access (only if NOT a local)
                if !is_local && self.is_class_field(name) {
                    if self.emit_self_ref() {
                        let field_name = self.canon(name);
                        let idx = self.str_const(&field_name);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                        return Ok(());
                    }
                }

                // Bare enum member: `Green` → `TColor.Green`
                if !is_local {
                    let canon_name = self.canon(name);
                    if let Some(enum_type) = self.enum_members.get(&canon_name).cloned() {
                        self.emit_global_read(&enum_type);
                        let mem_idx = self.str_const(&canon_name);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, mem_idx);
                        return Ok(());
                    }
                }

                // Bare profile namespace constant (e.g. Pascal `MaxInt`, `Pi`)
                if !is_local && !self.defined_globals.contains(&self.canon(name)) {
                    if let Some(cv) = self.profile.lookup_constant(name) {
                        match cv {
                            ConstantValue::Bool(b) => self.emit_const(Value::Bool(*b)),
                            ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                            ConstantValue::Str(s) => {
                                self.emit_const(Value::String(Arc::from(s.as_str())))
                            }
                        }
                        return Ok(());
                    }
                }

                // A bare reference to a parameterless function INVOKES it.
                // Pascal had its own copy of this behind a `profile.name ==
                // "pascal"` check; the only reason it needed one is that
                // `defined_functions` is INSERTED canonicalized (`link.rs`
                // `canon(name)`) but was looked up RAW here — so in a
                // case-insensitive language `Foo` never matched the stored
                // `foo`. `canon` reads `case_sensitive` off the profile, so
                // folding here is language-neutral and fixes VB too.
                //
                // The arity guard is pascal's, and it is right for both: a bare
                // name may only auto-invoke when the function takes no
                // parameters. `map_or(true, …)` keeps languages that do not
                // populate `function_min_arity` behaving as before.
                let canon_name = self.canon(name);
                if self.profile.bare_name_invokes_parameterless_function
                    && !is_local
                    && self.defined_functions.contains(&canon_name)
                    && self
                        .function_min_arity
                        .get(&canon_name)
                        .is_none_or(|arity| *arity == 0)
                {
                    self.emit_var_get(name);
                    self.emit_u8(Op::CALL_REF, 0);
                    return Ok(());
                }

                // ESM named import of a constant Value (e.g. `ecma:math::PI`).
                // These are NOT callable — they're inlined as compile-time
                // constants from `host_const_bindings` populated during
                // `collect_host_imports`. Function imports stay in
                // `host_import_bindings` and route through CALL_IMPORT.
                if !is_local {
                    let cname = self.canon(name);
                    if let Some(val) = self.host_const_bindings.get(&cname).cloned() {
                        self.emit_const(val);
                        return Ok(());
                    }
                }

                self.emit_var_get(name);
            }

            // ── This / Super ────────────────────────────────────────────
            ExprKind::This => {
                if self.profile.ambient_this_binding
                    && self.current_class.is_some()
                    && self.current_func_name.as_deref() != Some("<lambda>")
                    && self.current_func_name.as_deref().is_some_and(|name| {
                        !name.eq_ignore_ascii_case(&self.profile.constructor_name)
                    })
                {
                    self.emit_global_read("__js_this");
                    return Ok(());
                }

                let self_kw = &self.profile.self_keyword;
                if let Some(slot) = self
                    .scope()
                    .resolve(self_kw)
                    .or_else(|| self.scope().resolve("Self"))
                    .or_else(|| self.scope().resolve("self"))
                    .or_else(|| self.scope().resolve("this"))
                {
                    // §9.1.1.3.4 (JS): inside a derived constructor `this`
                    // is in TDZ until super() runs — reading it while
                    // this_slot is still null throws a ReferenceError.
                    if self.js_derived_ctor_ctx == Some((self.current, slot)) {
                        let l = self.line;
                        crate::primitives::classes::emit_this_initialized_guard(
                            self.chunk(),
                            slot,
                            l,
                        );
                    }
                    self.emit_u16(Op::LOCAL_GET, slot);
                } else if self.scopes.len() > 1 {
                    // Arrow function: capture `this` from enclosing scope via upvalue
                    let kw = self.profile.self_keyword.clone();
                    if let Some(_uv) = self.resolve_upvalue(self.scopes.len() - 1, &kw) {
                        let env = self.closure_env_slot();
                        let idx = self.closure_env_index(&kw);
                        let l = self.line;
                        crate::primitives::closures::emit_env_get(self.chunk(), env, idx, l);
                    } else if self.profile.ambient_this_binding {
                        if let Some(_uv) = self.resolve_upvalue(self.scopes.len() - 1, "__js_this")
                        {
                            let env = self.closure_env_slot();
                            let idx = self.closure_env_index("__js_this");
                            let l = self.line;
                            crate::primitives::closures::emit_env_get(self.chunk(), env, idx, l);
                        } else {
                            self.emit_global_read("__js_this");
                        }
                    } else if self.profile.ambient_this_binding {
                        self.emit_global_read("__js_this");
                    } else {
                        self.emit_null();
                    }
                } else if self.profile.ambient_this_binding {
                    self.emit_global_read("__js_this");
                } else {
                    self.emit_null();
                }
            }

            ExprKind::Super => {
                // super refers to the parent class constructor.
                // Look up the parent from the current class's PendingClass info.
                if let Some(ref class_name) = self.current_class.clone() {
                    if let Some(parent_name) = self
                        .pending_classes
                        .get(class_name.as_str())
                        .and_then(|pc| pc.parent.clone())
                    {
                        let pname = self.canon(&parent_name);
                        self.emit_global_read(&pname);
                    } else {
                        self.emit_null();
                    }
                } else {
                    self.emit_null();
                }
            }

            // ── Binary ──────────────────────────────────────────────────
            ExprKind::Binary { op, left, right } => {
                // Whether the expression is an INTEGER, per the language's own
                // `[builtin_types] int = [...]` spellings then the platform
                // table — builtinslotplan.md step 4a.
                //
                // This replaced TWO per-language spelling tables that lived
                // here: an exact 8-way match for C# and a `contains`-based one
                // for C (which also counts `char` and `bool`, correctly — they
                // ARE integer types in C). Both were shared-crate tables keyed
                // to one language, and the C one was reached through a literal
                let expr_is_integral = |compiler: &Compiler, expr: &Expression| {
                    matches!(expr.kind, ExprKind::Lit(Literal::Int(_)))
                        || compiler
                            .infer_expr_type_hint(expr)
                            .as_deref()
                            .is_some_and(|hint| {
                                vybe_ast::builtin_types::classify_with(
                                    &compiler.profile.builtin_type_spellings,
                                    hint,
                                ) == Some(vybe_ast::builtin_slots::BuiltinType::Int)
                            })
                };

                // Short-circuit for And/Or — generic path for all languages.
                // PHP used a custom truthiness path here that referenced the
                // removed __keys/vybe$assoc_keys_csv side-band and left the
                // stack unbalanced. emit_dyn_to_bool is correct: empty()/isset()
                // already handle PHP array truthiness at the call site.
                if matches!(op, BinOp::And | BinOp::Or)
                    && self.expr_is_integer_like(left)
                    && self.expr_is_integer_like(right)
                {
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    if *op == BinOp::And {
                        self.compile_binop(&BinOp::BitAnd);
                    } else {
                        self.compile_binop(&BinOp::BitOr);
                    }
                    return Ok(());
                }
                if *op == BinOp::And {
                    self.compile_expr(left)?;
                    let line = self.line;
                    let skip =
                        common::expressions::emit_and_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(
                        &mut self.chunks[self.current],
                        skip,
                    );
                    // A short circuit yields the OPERAND, which is a raw i32
                    // wherever it came from a comparison. `materialize_bool_results`
                    // is the profile property for exactly this — "a boolean is a
                    // VALUE in this language" — and it was already read at fifteen
                    // other operator sites; only `&&`/`||` still asked for Pascal by
                    // name. Pascal's own profile declares the property, so the name
                    // check could not change its answer, and every other language
                    // that declares it (C#, Go, Dart, Kotlin) declared it to stop
                    // `print(a == b)` printing `1` — which `print(a && b)` was still
                    // doing.
                    if self.profile.materialize_bool_results {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                    return Ok(());
                }
                if *op == BinOp::Or {
                    self.compile_expr(left)?;
                    let line = self.line;
                    let skip =
                        common::expressions::emit_or_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(
                        &mut self.chunks[self.current],
                        skip,
                    );
                    // Same as `&&` above — see the note there.
                    if self.profile.materialize_bool_results {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                    return Ok(());
                }
                // NullCoalesce as binary op
                if *op == BinOp::NullCoalesce {
                    self.compile_expr(left)?;
                    let value_slot = self.define_local("__null_coalesce_left");
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.compile_expr(right)?;
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_end(line);
                    return Ok(());
                }
                // Pow → BigInt fast path, then canonical stdlib path.
                if *op == BinOp::Pow {
                    let line = self.line;
                    if self.bigint_semantics()
                        && self.hint_is_bigint(self.infer_expr_type_hint(left).as_deref())
                        && self.hint_is_bigint(self.infer_expr_type_hint(right).as_deref())
                    {
                        let idx = self.import("ecma:bigint", "pow");
                        self.compile_expr(left)?;
                        self.compile_expr(right)?;
                        self.emit_host_call(idx, 2);
                        return Ok(());
                    }
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    common::math::emit_pow(self.chunk(), line);
                    return Ok(());
                }
                // InstanceOf → WASM GC `ref.test` opcode with the type name
                // from the const pool. The static (Ident RHS) form covers
                // every real-world `a instanceof TypeName` usage and matches
                // both GC `ref.test ht` and Component Model resource handle
                // typing — both of which require a compile-time-known type.
                // Dynamic RHS (the rare `a instanceof someVariable` JS form)
                // falls through to the host-less polyfill in compile_binop.
                if *op == BinOp::InstanceOf {
                    if self.class_prototype_dispatch() {
                        self.compile_expr(left)?;
                        match &right.kind {
                            // For built-in constructors that stamp `__type` on their instances,
                            // pass the name as a string — `js_instanceof` matches it against
                            // `__type` / `__types`, avoiding namespace-alias objects that have
                            // no `.name` property and would always yield `false`.
                            ExprKind::Ident(name)
                                if matches!(
                                    name.as_str(),
                                    "Error"
                                        | "EvalError"
                                        | "RangeError"
                                        | "ReferenceError"
                                        | "SyntaxError"
                                        | "TypeError"
                                        | "URIError"
                                        | "AggregateError"
                                        | "SuppressedError"
                                        | "RegExp"
                                        | "Map"
                                        | "Set"
                                        | "WeakMap"
                                        | "WeakSet"
                                        | "WeakRef"
                                        | "FinalizationRegistry"
                                        | "ArrayBuffer"
                                        | "DataView"
                                        | "Date"
                                        | "Promise"
                                        | "SharedArrayBuffer"
                                        | "URL"
                                        | "URLSearchParams"
                                        | "TextEncoder"
                                        | "TextDecoder"
                                ) =>
                            {
                                self.emit_const(Value::String(Arc::from(name.as_str())));
                            }
                            _ => self.compile_expr(right)? }
                        self.compile_binop(op);
                        return Ok(());
                    }
                    if let crate::ast::ExprKind::Ident(type_name) = &right.kind {
                        self.compile_expr(left)?;
                        let line = self.line;
                        // A source type alias (`use App\Http\Request;`)
                        // resolves to the declared namespace-qualified
                        // identity, so `x instanceof Request` REF_TESTs
                        // the real `app.http.request` type.
                        let aliased = self.resolve_source_type_alias(type_name);
                        let name_canon = self.canon(&aliased);
                        // Through the shared reflection primitive so `instanceof`
                        // and a typed `catch` cannot disagree about the same
                        // object — it spills the operand because the answer needs
                        // to read it twice (rtt, then the `__types` chain).
                        let obj_slot = self.define_local("__instanceof_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);
                        crate::primitives::reflection::emit_is_instance_of(
                            &mut self.chunks,
                            self.current,
                            obj_slot,
                            &name_canon,
                            line,
                        );
                        return Ok(());
                    }
                }
                if self.profile.ecma_in_operator && *op == BinOp::In {
                    if let ExprKind::Ident(field) = &left.kind {
                        if field.starts_with('#') {
                            let storage_name = self.js_member_storage_name(field);
                            self.emit_const(Value::String(Arc::from(storage_name.as_str())));
                            self.compile_expr(right)?;
                            self.compile_binop(op);
                            return Ok(());
                        }
                    }
                }
                // A language whose `in` means VALUE membership declares
                // `[builtin_slots.array] contains`; one that means a KEY test
                // (the ECMA profiles, `ecma:object.hasOwn`) declares nothing and
                // cannot reach this. When the receiver IS a set, fall through to
                // `compile_binop`, which uses the `set` row instead.
                if (*op == BinOp::In || *op == BinOp::NotIn) && !self.expr_is_builtin_set(right) {
                    if let Some(target) = self.array_contains_target() {
                        self.compile_expr(left)?;
                        self.compile_expr(right)?;
                        self.emit_contains(&target);
                        if *op == BinOp::NotIn {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                        }
                        return Ok(());
                    }
                }
                if self.try_compile_set_arithmetic(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_python_set_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_csharp_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_dotnet_datetime_timespan_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_fortran_interface_binary_operator(op, left, right)? {
                    return Ok(());
                }
                if self.try_compile_fortran_array_binary_operator(op, left, right)? {
                    return Ok(());
                }

                // `int / int` truncates. ONE branch for every language that
                // says so: C now declares `integer_division_on_slash = true`
                // like C# always did, and each brings its own integer
                // spellings, so the second copy of this block — the one gated
                if self.profile.integer_division_on_slash
                    && *op == BinOp::Div
                    && expr_is_integral(self, left)
                    && expr_is_integral(self, right)
                {
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    self.compile_binop(&BinOp::IDiv);
                    return Ok(());
                }

                // `xor` resolves by operand type — bitwise on integers,
                if self.profile.xor_is_logical_for_non_integers && *op == BinOp::BitXor {
                    let integer_xor = self.expr_is_integer_like(left)
                        && self.expr_is_integer_like(right);
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    self.compile_binop(&BinOp::BitXor);
                    if !integer_xor {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                    return Ok(());
                }

                // PHP `.` and Lua `..` coerce BOTH operands to string up
                // front instead of relying on the concat op's own coercion.
                // The coercion itself is the only difference between them —
                // PHP renders `true` as `"1"`, `null` as `""` and an array as
                // `"Array"`, none of which is ECMA `String()` — so the
                // language supplies it through `concat_stringify` and the
                // shared `to_string` is the default.
                if self.profile.concat_stringifies_operands && *op == BinOp::Concat {
                    let line = self.line;
                    // `[builtin_slots.string] to_string` — was
                    // `LanguageHooks::concat_stringify`, looked up by language
                    // NAME (builtinslotplan.md §3c). Undeclared falls back to
                    // the shared ECMA `String()` coercion, so a language that
                    // declares nothing is unaffected.
                    let stringify = self
                        .profile
                        .builtin_slots
                        .get(
                            vybe_ast::builtin_slots::BuiltinType::String,
                            vybe_ast::ProtocolSlot::ToString,
                        )
                        .map(str::to_string);
                    for operand in [left, right] {
                        self.compile_expr(operand)?;
                        match &stringify {
                            Some(target) => {
                                self.emit_slot_target(target, 1, line, "string to_string")
                            }
                            None => common::strings::emit_to_string(self.chunk(), line) }
                    }
                    self.compile_binop(op);
                    return Ok(());
                }

                // BigInt arithmetic and comparisons via ecma:bigint host fns.
                // These already exist and return Value::BigInt / Value::Bool.
                // `infer_expr_type_hint` returns "bigint" for BigInt literals
                // and for variables initialised with BigInt values.
                if self.bigint_semantics() {
                    let left_hint = self.infer_expr_type_hint(left);
                    let right_hint = self.infer_expr_type_hint(right);
                    let left_is_bigint = self.hint_is_bigint(left_hint.as_deref());
                    let right_is_bigint = self.hint_is_bigint(right_hint.as_deref());
                    // The other operand is "known non-BigInt" only when its
                    // type was inferred to something concrete that isn't
                    // bigint. An *unknown* hint (e.g. a reassigned parameter
                    // holding a BigInt at runtime) is NOT a compile-time mix
                    // error — §21.2.1.1 mixing is a RUNTIME TypeError, so we
                    // route to the bigint op and let it coerce/throw.
                    let other_known_non_bigint = if left_is_bigint {
                        right_hint.is_some()
                    } else {
                        left_hint.is_some()
                    };
                    if left_is_bigint && right_is_bigint {
                        let fn_name: Option<&str> = match op {
                            BinOp::Add => Some("add"),
                            BinOp::Sub => Some("sub"),
                            BinOp::Mul => Some("mul"),
                            BinOp::Div => Some("div"),
                            BinOp::Mod => Some("rem"),
                            BinOp::BitAnd => Some("and"),
                            BinOp::BitOr => Some("or"),
                            BinOp::BitXor => Some("xor"),
                            BinOp::Pow => Some("pow"),
                            BinOp::Shl => Some("shl"),
                            BinOp::Shr => Some("shr"),
                            BinOp::Eq | BinOp::StrictEq => Some("eq"),
                            BinOp::NotEq | BinOp::StrictNotEq => Some("ne"),
                            BinOp::Lt => Some("lt"),
                            BinOp::LtEq => Some("le"),
                            BinOp::Gt => Some("gt"),
                            BinOp::GtEq => Some("ge"),
                            _ => None };
                        if let Some(name) = fn_name {
                            let idx = self.import("ecma:bigint", name);
                            self.compile_expr(left)?;
                            self.compile_expr(right)?;
                            self.emit_host_call(idx, 2);
                            return Ok(());
                        }
                    } else if left_is_bigint || right_is_bigint {
                        let arith_fn: Option<&str> = match op {
                            BinOp::Add => Some("add"),
                            BinOp::Sub => Some("sub"),
                            BinOp::Mul => Some("mul"),
                            BinOp::Div => Some("div"),
                            BinOp::Mod => Some("rem"),
                            BinOp::BitAnd => Some("and"),
                            BinOp::BitOr => Some("or"),
                            BinOp::BitXor => Some("xor"),
                            BinOp::Shl => Some("shl"),
                            BinOp::Shr => Some("shr"),
                            BinOp::Pow => Some("pow"),
                            _ => None };
                        if let Some(name) = arith_fn {
                            if other_known_non_bigint {
                                // A language that DECLARES integer spellings
                                // as bigint widens its own mixes (Kotlin
                                // `1L + 2` is Long, `1L + 2.0` is Double):
                                // with a floating partner the BIGINT side
                                // demotes through `Number`; otherwise the
                                // number side promotes through the ecma
                                // constructor and the op stays exact.
                                if self.bigint_widens_mixes() {
                                    let other_hint = if left_is_bigint {
                                        right_hint.as_deref()
                                    } else {
                                        left_hint.as_deref()
                                    };
                                    let other_is_float = other_hint.is_some_and(|h| {
                                        vybe_ast::builtin_types::classify_with(
                                            &self.profile.builtin_type_spellings,
                                            h,
                                        ) == Some(
                                            vybe_ast::builtin_slots::BuiltinType::Double,
                                        )
                                    });
                                    if other_is_float {
                                        let number = self.import("ecma:number", "Number");
                                        self.compile_expr(left)?;
                                        if left_is_bigint {
                                            self.emit_host_call(number, 1);
                                        }
                                        self.compile_expr(right)?;
                                        if right_is_bigint {
                                            self.emit_host_call(number, 1);
                                        }
                                        self.compile_binop(op);
                                        return Ok(());
                                    }
                                    let ctor = self.import("ecma:bigint", "BigInt");
                                    let idx = self.import("ecma:bigint", name);
                                    self.compile_expr(left)?;
                                    if !left_is_bigint {
                                        self.emit_host_call(ctor, 1);
                                    }
                                    self.compile_expr(right)?;
                                    if !right_is_bigint {
                                        self.emit_host_call(ctor, 1);
                                    }
                                    self.emit_host_call(idx, 2);
                                    return Ok(());
                                }
                                // §21.2.1.1: arithmetic between BigInt and a
                                // KNOWN non-BigInt throws TypeError.
                                self.emit_const(Value::String(Arc::from(
                                    "Cannot mix BigInt and other types, use explicit conversions",
                                )));
                                let line = self.line;
                                self.emit_js_exception_ctor_from_message_value("TypeError")?;
                                common::errors::emit_throw(self.chunk(), line);
                                return Ok(());
                            }
                            // Other operand is unknown — route to the bigint
                            // op (its `to_bigint` coerces an actual BigInt and
                            // throws at runtime on a real Number).
                            let idx = self.import("ecma:bigint", name);
                            self.compile_expr(left)?;
                            self.compile_expr(right)?;
                            self.emit_host_call(idx, 2);
                            return Ok(());
                        } else if matches!(
                            op,
                            BinOp::Eq
                                | BinOp::NotEq
                                | BinOp::Lt
                                | BinOp::LtEq
                                | BinOp::Gt
                                | BinOp::GtEq
                        ) {
                            // §7.2.13: BigInt == Number compares numerically.
                            // ecma:bigint comparisons use to_bigint() on both sides, which handles
                            // all numeric types (F64, I32, BigInt) correctly.
                            let fn_name = match op {
                                BinOp::Eq => "eq",
                                BinOp::NotEq => "ne",
                                BinOp::Lt => "lt",
                                BinOp::LtEq => "le",
                                BinOp::Gt => "gt",
                                BinOp::GtEq => "ge",
                                _ => unreachable!() };
                            let idx = self.import("ecma:bigint", fn_name);
                            self.compile_expr(left)?;
                            self.compile_expr(right)?;
                            self.emit_host_call(idx, 2);
                            return Ok(());
                        }
                        // StrictEq / StrictNotEq with mixed types: fall through to compile_binop
                        // (strict equality of different types returns false/true without coercion).
                    }
                }

                self.compile_expr(left)?;
                self.compile_expr(right)?;
                self.compile_binop(op);
                if self.profile.materialize_bool_results
                    && matches!(
                        op,
                        BinOp::Eq
                            | BinOp::NotEq
                            | BinOp::StrictEq
                            | BinOp::StrictNotEq
                            | BinOp::Lt
                            | BinOp::LtEq
                            | BinOp::Gt
                            | BinOp::GtEq
                            | BinOp::In
                            | BinOp::NotIn
                    )
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                }
            }

            // ── Unary ───────────────────────────────────────────────────
            ExprKind::Unary { op, expr: inner } => {
                match op {
                    UnaryOp::PreInc | UnaryOp::PostInc => {
                        // ++x / x++ : load, add 1, store
                        self.compile_expr(inner)?;
                        if *op == UnaryOp::PostInc {
                            inst!(self, core_wasm::dup);
                        }
                        self.emit_step_by_one(true);
                        if *op == UnaryOp::PreInc {
                            inst!(self, core_wasm::dup);
                        }
                        self.compile_assign_target(inner)?;
                    }
                    UnaryOp::PreDec | UnaryOp::PostDec => {
                        self.compile_expr(inner)?;
                        if *op == UnaryOp::PostDec {
                            inst!(self, core_wasm::dup);
                        }
                        self.emit_step_by_one(false);
                        if *op == UnaryOp::PreDec {
                            inst!(self, core_wasm::dup);
                        }
                        self.compile_assign_target(inner)?;
                    }
                    UnaryOp::AddrOf => {
                        self.compile_address_of_expr(inner)?;
                    }
                    UnaryOp::Deref => {
                        self.compile_deref_expr(inner)?;
                    }
                    _ => {
                        if self.profile.ecma_typeof_operator && matches!(op, UnaryOp::Typeof) {
                            if let ExprKind::Ident(name) = &inner.kind {
                                let is_callable_global =
                                    matches!(
                                        name.as_str(),
                                        "eval"
                                            | "parseInt"
                                            | "parseFloat"
                                            | "Function"
                                            | "Object"
                                            | "Boolean"
                                            | "Number"
                                            | "String"
                                            | "Array"
                                            | "Symbol"
                                            | "BigInt"
                                            | "Date"
                                            | "RegExp"
                                            | "Promise"
                                            | "Proxy"
                                            | "Map"
                                            | "Set"
                                            | "WeakMap"
                                            | "WeakSet"
                                            | "ArrayBuffer"
                                            | "SharedArrayBuffer"
                                            | "DataView"
                                            | "Int8Array"
                                            | "Uint8Array"
                                            | "Uint8ClampedArray"
                                            | "Int16Array"
                                            | "Uint16Array"
                                            | "Int32Array"
                                            | "Uint32Array"
                                            | "Float32Array"
                                            | "Float64Array"
                                            | "BigInt64Array"
                                            | "BigUint64Array"
                                    ) || self.profile.lookup_builtin(name).is_some();
                                if is_callable_global {
                                    self.emit_const(Value::String(Arc::from("function")));
                                    return Ok(());
                                }
                            }
                        }
                        self.compile_expr(inner)?;
                        match op {
                            UnaryOp::Neg => {
                                let l = self.line;
                                if self.bigint_semantics()
                                    && self.hint_is_bigint(self.infer_expr_type_hint(inner).as_deref())
                                {
                                    let idx = self.import("ecma:bigint", "neg");
                                    self.emit_host_call(idx, 1);
                                } else if self.uses_rich_operators() {
                                    // A user `operator -()` / `__neg__` defines
                                    // negation for its own type; numeric
                                    // negation would coerce it instead.
                                    self.emit_rich_unary(
                                        vybe_ast::ProtocolSlot::Neg,
                                        common::math::emit_neg,
                                    );
                                } else {
                                    common::math::emit_neg(self.chunk(), l);
                                }
                            }
                            UnaryOp::Pos => {
                                // JS `+v` coerces to number — ECMA-262 §7.1.4 ToNumber.
                                // BigInt is the one primitive exception: unary
                                // plus throws, while explicit Number(1n) still
                                // converts.
                                if self.bigint_semantics()
                                    && self.hint_is_bigint(self.infer_expr_type_hint(inner).as_deref())
                                {
                                    self.emit_const(Value::String(Arc::from(
                                        "Cannot convert a BigInt value to a number",
                                    )));
                                    let line = self.line;
                                    self.emit_js_exception_ctor_from_message_value("TypeError")?;
                                    common::errors::emit_throw(self.chunk(), line);
                                    return Ok(());
                                }
                                // For Object operands, ToPrimitive(hint=number)
                                // first (so Symbol.toPrimitive / valueOf
                                // overrides fire with the JS method-call
                                // protocol intact), then `Number(...)` to
                                // finalise the coercion.
                                if self.profile.ecma_to_primitive {
                                    self.emit_to_primitive("number");
                                }
                                let idx = self.import("ecma:number", "Number");
                                self.emit_host_call(idx, 1);
                            }
                            UnaryOp::Not => {
                                let line = self.line;
                                if self.expr_is_integer_like(inner)
                                {
                                    common::expressions::emit_i32_not(self.chunk(), line);
                                } else {
                                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                                }
                                // `expr_is_integer_like` already answers false
                                // unless the language declares
                                // `logical_ops_bitwise_for_integers`, so no
                                // language check is needed here.
                                if (self.profile.ecma_boolean_operators
                                    || self.profile.materialize_bool_results)
                                    && !self.expr_is_integer_like(inner)
                                {
                                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                                }
                            }
                            UnaryOp::BitNot => {
                                let l = self.line;
                                // ECMA §13.5.6: `~` dispatches on ToNumeric — a
                                // BigInt operand takes the `(bigint, not)` slot
                                // row, anything else ToInt32. The spelling lives
                                // in the slot table (language first, platform
                                // default second), not here.
                                let bigint_not = self
                                    .bigint_semantics()
                                    .then(|| {
                                        self.builtin_type_slot_target(
                                            vybe_ast::builtin_slots::BuiltinType::BigInt,
                                            vybe_ast::ProtocolSlot::Not,
                                        )
                                    })
                                    .flatten()
                                    .map(str::to_string);
                                if let Some(target) = bigint_not {
                                    if self.hint_is_bigint(
                                        self.infer_expr_type_hint(inner).as_deref(),
                                    ) {
                                        // Statically known BigInt.
                                        self.emit_slot_target(&target, 1, l, "bigint not");
                                    } else {
                                        // Unknown operand: runtime ToNumeric
                                        // dispatch, the same shape as the
                                        // `emit_dyn_*` family.
                                        let slot = self.chunk().alloc_scratch(1);
                                        self.emit_u16(Op::LOCAL_SET, slot);
                                        let test_bi = self.import("wasm:js-bigint", "test");
                                        self.emit_u16(Op::LOCAL_GET, slot);
                                        self.emit_host_call(test_bi, 1);
                                        self.chunk().emit_if_value(l);
                                        self.emit_u16(Op::LOCAL_GET, slot);
                                        self.emit_slot_target(&target, 1, l, "bigint not");
                                        self.chunk().emit_else(l);
                                        self.emit_u16(Op::LOCAL_GET, slot);
                                        common::expressions::emit_i32_not(self.chunk(), l);
                                        self.chunk().emit_end(l);
                                    }
                                } else if self.uses_rich_operators() {
                                    self.emit_rich_unary(
                                        vybe_ast::ProtocolSlot::Not,
                                        common::expressions::emit_i32_not,
                                    );
                                } else {
                                    common::expressions::emit_i32_not(self.chunk(), l);
                                }
                            }
                            UnaryOp::Typeof => fn_call!(self, "ecma:value", "typeof", 1),
                            UnaryOp::Void => {
                                self.emit(Op::DROP);
                                inst!(self, core_wasm::undefined);
                            }
                            UnaryOp::Delete => {
                                self.emit(Op::DROP);
                                inst!(self, core_wasm::bool_const, true);
                            }
                            _ => {}              // PreInc etc handled above
                        }
                    }
                }
            }

            ExprKind::RefOf(place) => match place.as_ref() {
                PlaceExpr::Ident(name) => {
                    if let Some(slot) = self.promote_local_binding_to_pointer_cell(name) {
                        self.emit_u16(Op::LOCAL_GET, slot);
                    } else {
                        self.compile_expr(&Expression::ident(name))?;
                        self.emit_wrap_top_of_stack_in_pointer_cell();
                    }
                }
                PlaceExpr::Deref(expr) => {
                    self.compile_expr(expr)?;
                }
                PlaceExpr::Member { object, field, .. } => {
                    self.compile_member_reference(object, field)?;
                }
                // A PLACE must never reach the rvalue wrap — see
                // `compile_index_reference`. Both spellings of "address of an
                // element" resolve through the one site.
                PlaceExpr::Index { object, index, .. } => {
                    self.compile_index_reference(object, index)?;
                }
            },

            ExprKind::RefLoad(expr) => {
                self.compile_deref_expr(expr)?;
            }

            // ── Ternary ─────────────────────────────────────────────────
            ExprKind::Ternary { cond, then, else_ } => {
                self.compile_expr(cond)?;
                self.emit_condition_truthiness_from_stack();
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.compile_expr(then)?;
                self.chunk().emit_else(line);
                self.compile_expr(else_)?;
                self.chunk().emit_end(line);
            }

            // ── Call ────────────────────────────────────────────────────
            ExprKind::Call {
                callee,
                args,
                optional } => {
                if *optional {
                    // Optional call: callee?.() — short-circuit to undefined if callee is null/undefined.
                    // Per ECMA-262 §13.5.9: the result is `undefined` (not null) when short-circuiting.
                    self.compile_expr(callee)?;
                    let tmp = self.define_local("__optional_callee");
                    self.emit_u16(Op::LOCAL_SET, tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    inst!(self, core_wasm::undefined);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    let undef_line = self.line;
                    self.chunk().emit_if_value(undef_line);
                    inst!(self, core_wasm::undefined);
                    self.chunk().emit_else(undef_line);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    for a in args {
                        self.compile_expr(&a.value)?;
                    }
                    self.emit_u8(Op::CALL_REF, args.len() as u8);
                    self.chunk().emit_end(undef_line);
                    self.chunk().emit_end(line);
                } else {
                    if self.profile.parens_for_index
                        && !args.is_empty()
                        && matches!(&callee.kind,
                            ExprKind::Ident(name) if self.lookup_array_binding(name).is_some()
                        )
                    {
                        self.compile_expr(callee)?;
                        for arg in args {
                            self.compile_array_index_operand_for_owner(callee, &arg.value)?;
                            let line = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, line);
                        }
                        return Ok(());
                    }

                    self.compile_call(callee, args)?;
                    // Multi-value result repack: when the callee is one of
                    // the pre-scanned multi-return functions, CALL leaves
                    // N values on the stack. A destructure-assign consumes
                    // them directly (see `detect_multi_value_receive` in
                    // primitives/mod.rs, which bypasses this branch); every
                    // other use site — `r = f()`, `print(f())`, `f() + g()`
                    // — expects a single value, so we re-pack here.
                    if let Some(n) = self.multi_return_arity_for_callee(callee) {
                        self.pack_multi_value_result(n);
                    }
                }
            }

            // ── Member access ───────────────────────────────────────────
            ExprKind::Member {
                object,
                field,
                null_safe } => {
                if self.private_member_access_forbidden(field) {
                    self.emit_private_access_denied(field)?;
                    return Ok(());
                }
                let source_member_parts = self.flatten_member_chain(expr);
                if source_member_parts.len() >= 2 {
                    if let Some(source_function) =
                        self.resolve_namespaced_function_identity(&source_member_parts.join("."))
                    {
                        self.emit_global_read(&source_function);
                        return Ok(());
                    }
                }
                // `Task<T>.Result` blocks until the task completes — a dotnet
                // trait, not a VB one. Scoped to `name == "vb"` it left C#
                // printing nothing for `t.Result`. `canon` matches the way the
                // profile says identifiers match, so a case-sensitive dotnet
                // language does not also catch a member named `result`.
                if self.profile.namespaces.use_dotnet && self.canon(field) == self.canon("Result")
                {
                    let obj_slot = self.define_local("__dotnet_task_result_obj");
                    let value_slot = self.define_local("__dotnet_task_result_value");
                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let result_key = self.str_const("Result");
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, result_key);
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let undef_idx = self.import("wasm:js-undefined", "test");
                    self.emit_host_call(undef_idx, 1);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    crate::primitives::functions::emit_await(self.chunk(), line);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_end(line);
                    return Ok(());
                }
                // The COMMON resolver, asked without a language gate. It walks
                // the profile's own declared namespace surface
                // (`host_namespace_aliases`, `host_package_roots`, the shared
                // tree scoped by `type_scopes`), so a profile that declares none
                // resolves nothing and this falls straight through. The gate
                // decided who was allowed to ask, never what came back.
                {
                    let parts = self.flatten_member_chain(expr);
                    if let Some(super::resolver::Resolution::ResolvedPrefix { target, suffix }) =
                        self.resolve_profile_namespace_chain(&parts)
                    {
                        match target {
                            crate::primitives::namespaces::ResolutionTarget::CommonEmit(emit) => {
                                self.emit_common(&emit, 0, self.line);
                            }
                            crate::primitives::namespaces::ResolutionTarget::HostCall {
                                module,
                                func,
                                ..
                            } => {
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, 0);
                            }
                            crate::primitives::namespaces::ResolutionTarget::Const(value) => {
                                self.emit_const(value);
                            }
                            _ => {}
                        }
                        for part in suffix {
                            if part.eq_ignore_ascii_case("Length")
                                || part.eq_ignore_ascii_case("Count")
                            {
                                common::collections::emit_len(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                            } else {
                                let idx = self.str_const(&part);
                                self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                            }
                        }
                        return Ok(());
                    }
                }
                // First-class function reference (WASM `ref.func $f`): with
                // `function_references`, a static method named as a VALUE (not
                // called) tears off into a `REF_FUNC` funcref rather than being
                // invoked or read as a bound-method property. Gated on the
                // profile property so no other language's member semantics
                // change. `object` must be a known class (not a local holding
                // an object) and `field` must resolve to a single function
                // chunk.
                if self.profile.function_references && !*null_safe {
                    if let ExprKind::Ident(obj_name) = &object.kind {
                        let class = self.canon(obj_name);
                        // The class may not be in `defined_classes` yet when a
                        // pre-`_start` init (e.g. an `elem` segment) references
                        // it — `pending_classes` (populated in the pre-pass) is
                        // the authoritative "is a known class" signal here.
                        if (self.defined_classes.contains(&class)
                            || self.pending_classes.contains_key(&class))
                            && self.scope().resolve(obj_name).is_none()
                        {
                            if let Some(chunk_idx) =
                                self.resolve_unique_static_method_chunk_for_class(&class, field)
                            {
                                let line = self.line;
                                self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
                                self.chunk().emit(0, line); // uv_count = 0
                                return Ok(());
                            }
                        }
                    }
                }
                // §19.3 global-object semantics: `globalThis.X` reads resolve
                // the global X — the twin of the write special-case (Assign:
                // `globalThis.X = v` also GLOBAL_SETs X). The shared globalThis
                // singleton object does not carry host-installed builtins, so a
                // plain property read would yield undefined for e.g.
                // `globalThis.Map` — which silently disabled the prelude's
                // Map/Set/Boolean/Number wrappers.
                // BISECT: temporarily disabled — see below.
                // BISECT left in place: this read path is disabled in the
                // tree and re-enabling it is not part of unifying the spelling.
                if false
                    && self.profile.dynamic_member_access
                    && !*null_safe
                    && matches!(&object.kind, ExprKind::Ident(n)
                        if crate::primitives::globals::names_global_namespace(&self.profile, n))
                {
                    self.emit_var_get(field);
                    return Ok(());
                }
                let reflection_field = field.split('<').next().unwrap_or(field.as_str()).trim();
                if let Some(binding) = self.resolve_reflection_binding_expr(object) {
                    match (binding, reflection_field) {
                        (ReflectionBinding::Type(type_name), "Name") => {
                            let short_name = self.reflection_type_short_name(&type_name);
                            self.compile_expr(&Expression::string(&short_name))?;
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "FullName") => {
                            let full_name = self.reflection_type_full_name(&type_name);
                            self.compile_expr(&Expression::string(&full_name))?;
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "IsEnum") => {
                            let v = self.reflection_is_enum_type(&type_name);
                            inst!(self, core_wasm::bool_const, v);
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "IsValueType") => {
                            let v = self.reflection_is_value_type(&type_name);
                            inst!(self, core_wasm::bool_const, v);
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "IsSealed") => {
                            let v = self.reflection_is_sealed_type(&type_name);
                            inst!(self, core_wasm::bool_const, v);
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "IsSerializable") => {
                            let attrs = self.reflection_attributes_for_type(
                                &type_name,
                                Some("System.SerializableAttribute"),
                                true,
                            );
                            inst!(self, core_wasm::bool_const, !attrs.is_empty());
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "BaseType") => {
                            if let Some(parent_type) = self.reflection_base_type_name(&type_name) {
                                self.compile_reflection_type_value(&parent_type)?;
                            } else {
                                self.emit_null();
                            }
                            return Ok(());
                        }
                        (ReflectionBinding::Type(type_name), "DeclaringType") => {
                            if let Some(parent_type) =
                                self.reflection_declaring_type_name(&type_name)
                            {
                                self.compile_reflection_type_value(&parent_type)?;
                            } else {
                                self.emit_null();
                            }
                            return Ok(());
                        }
                        (ReflectionBinding::AssemblyName, "Name") => {
                            self.compile_expr(&Expression::string("main"))?;
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Method {
                                type_name,
                                method_name,
                                ..
                            },
                            "IsStatic",
                        ) => {
                            let is_static = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| meta.methods.get(&method_name))
                                .map(|meta| meta.is_static)
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, is_static);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Method {
                                type_name,
                                method_name,
                                generic_args },
                            "ReturnType",
                        ) => {
                            let return_type = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| meta.methods.get(&method_name))
                                .and_then(|meta| {
                                    meta.return_type.as_deref().map(|return_type| {
                                        if let Some(index) = meta
                                            .generic_params
                                            .iter()
                                            .position(|param| param.eq_ignore_ascii_case(return_type))
                                            .filter(|index| *index < generic_args.len())
                                        {
                                            generic_args[index].clone()
                                        } else {
                                            return_type.to_string()
                                        }
                                    })
                                })
                                .unwrap_or_else(|| "System.Void".to_string());
                            self.compile_reflection_type_value(&return_type)?;
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Method {
                                type_name,
                                method_name,
                                ..
                            },
                            "IsAbstract",
                        ) => {
                            let value = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| meta.methods.get(&method_name))
                                .map(|meta| meta.is_abstract)
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Method {
                                type_name,
                                method_name,
                                ..
                            },
                            "IsVirtual",
                        ) => {
                            let value = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| meta.methods.get(&method_name))
                                .map(|meta| meta.is_virtual)
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Method {
                                type_name,
                                method_name,
                                generic_args },
                            "IsGenericMethodDefinition",
                        ) => {
                            let value = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| meta.methods.get(&method_name))
                                .map(|meta| !meta.generic_params.is_empty() && generic_args.is_empty())
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Property {
                                type_name,
                                property_name },
                            "CanRead",
                        ) => {
                            let can_read = self
                                .reflection_property_metadata(&type_name, &property_name)
                                .is_some();
                            inst!(self, core_wasm::bool_const, can_read);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Property {
                                type_name,
                                property_name },
                            "CanWrite",
                        ) => {
                            let can_write = self
                                .reflection_property_metadata(&type_name, &property_name)
                                .map(|(_, meta)| meta)
                                .map(|meta| meta.can_write)
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, can_write);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Field {
                                type_name,
                                field_name },
                            "IsPublic",
                        ) => {
                            let value = self
                                .reflection_field_metadata(&type_name, &field_name)
                                .map(|(_, meta)| {
                                    matches!(meta.visibility, vybe_ast::Visibility::Public)
                                })
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Constructor {
                                type_name,
                                param_types },
                            "IsPublic",
                        ) => {
                            let value = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| {
                                    meta.constructors
                                        .iter()
                                        .find(|ctor| ctor.param_types == param_types)
                                })
                                .map(|meta| matches!(meta.visibility, vybe_ast::Visibility::Public))
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Constructor {
                                type_name,
                                param_types },
                            "IsPrivate",
                        ) => {
                            let value = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| {
                                    meta.constructors
                                        .iter()
                                        .find(|ctor| ctor.param_types == param_types)
                                })
                                .map(|meta| matches!(meta.visibility, vybe_ast::Visibility::Private))
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Constructor {
                                type_name,
                                param_types },
                            "IsStatic",
                        ) => {
                            let value = self
                                .reflection_type_metadata(&type_name)
                                .and_then(|meta| {
                                    meta.constructors
                                        .iter()
                                        .find(|ctor| ctor.param_types == param_types)
                                })
                                .map(|meta| meta.is_static)
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Field {
                                type_name,
                                field_name },
                            "IsPrivate",
                        ) => {
                            let value = self
                                .reflection_field_metadata(&type_name, &field_name)
                                .map(|(_, meta)| {
                                    matches!(meta.visibility, vybe_ast::Visibility::Private)
                                })
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Field {
                                type_name,
                                field_name },
                            "IsStatic",
                        ) => {
                            let value = self
                                .reflection_field_metadata(&type_name, &field_name)
                                .map(|(_, meta)| meta.is_static)
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        (
                            ReflectionBinding::Field {
                                type_name,
                                field_name },
                            "IsInitOnly",
                        ) => {
                            let value = self
                                .reflection_field_metadata(&type_name, &field_name)
                                .map(|(_, meta)| !meta.can_write)
                                .unwrap_or(false);
                            inst!(self, core_wasm::bool_const, value);
                            return Ok(());
                        }
                        _ => {}
                    }
                }

                if field.eq_ignore_ascii_case("IsEnum") {
                    if let ExprKind::Lit(Literal::Str(type_name)) = &object.kind {
                        let full_name = type_name
                            .strip_prefix("System.")
                            .unwrap_or(type_name.as_str())
                            .trim();
                        let short_name = full_name.rsplit('.').next().unwrap_or(full_name).trim();
                        let is_enum = [full_name, short_name].into_iter().any(|candidate| {
                            let canon = self.canon(candidate);
                            self.enum_value_names.contains_key(&canon)
                                || self
                                    .enum_value_names
                                    .keys()
                                    .any(|known| known.eq_ignore_ascii_case(candidate))
                        });
                        inst!(self, core_wasm::bool_const, is_enum);
                        return Ok(());
                    }
                }

                // Namespace constant check (Math.PI, etc.)
                if let ExprKind::Ident(obj_name) = &object.kind {
                    // An EXACT local match always wins. A match that only
                    // succeeds under case folding loses to a type-qualified
                    // interpretation — `Foo.Bar` naming a type beats a local
                    // spelled `foo`. This is the one place the exact/folded
                    // distinction is the actual question, hence `resolve_exact`.
                    let prefers_type_lookup =
                        self.prefers_type_qualified_member_lookup(obj_name, field);
                    let obj_is_local = self.scope().resolve_exact(obj_name).is_some()
                        || self.has_static_local_binding(obj_name)
                        || (self.scope().resolve(obj_name).is_some() && !prefers_type_lookup);
                    if !obj_is_local {
                        if let Some(value) = self.enum_member_ordinal(obj_name, field) {
                            self.emit_const(Value::F64(value as f64));
                            return Ok(());
                        }
                    }

                    if let Some(key) = self.generic_static_member_key(obj_name, field) {
                        self.emit_global_read(&key);
                        return Ok(());
                    }

                    let compound = format!("{}.{}", obj_name, field);
                    if let Some(cv) = self.profile.lookup_constant(&compound) {
                        match cv {
                            ConstantValue::Bool(b) => self.emit_const(Value::Bool(*b)),
                            ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                            ConstantValue::Str(s) => {
                                self.emit_const(Value::String(Arc::from(s.as_str())))
                            }
                        }
                        return Ok(());
                    }
                    // Constructor call with 0 args: ClassName.Create
                    let ctor_nm = &self.profile.constructor_name;
                    let is_ctor = if self.case_sensitive {
                        field == ctor_nm
                    } else {
                        field.eq_ignore_ascii_case(ctor_nm)
                    };
                    let canon_obj = self.canon(obj_name);
                    let is_known_class = self.defined_classes.contains(&canon_obj)
                        && self.scope().resolve(obj_name).is_none();
                    if is_ctor && is_known_class {
                        self.emit_var_get(obj_name);
                        self.emit_u8(Op::CALL_REF, 0);
                        return Ok(());
                    }

                    // A parameterless CLASS function used without `()`:
                    // `TShape.Circle`. Only auto-invoke when the member resolves
                    // to a known static method whose chunk arity is
                    // receiver-only (the class object plus zero user args).
                    if self.profile.member_invokes_parameterless_method {
                        let is_class = self.defined_classes.contains(&canon_obj)
                            && self.scope().resolve(obj_name).is_none();
                        if is_class {
                            let method_name = self.canon(field);
                            let zero_arg_static = self
                                .pending_classes
                                .get(canon_obj.as_str())
                                .map(|pc| !pc.static_fields.iter().any(|name| name == &method_name))
                                .unwrap_or(false);
                            if zero_arg_static {
                                self.emit_global_read(&canon_obj);
                                inst!(self, core_wasm::dup);
                                let method_idx = self.str_const(&method_name);
                                self.emit_struct_field_op(Op::STRUCT_GET, 0, method_idx);
                                let fn_tmp = self
                                    .scope()
                                    .resolve("__pascal_static_fn")
                                    .unwrap_or_else(|| self.define_local("__pascal_static_fn"));
                                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                                let cls_tmp = self
                                    .scope()
                                    .resolve("__pascal_static_cls")
                                    .unwrap_or_else(|| self.define_local("__pascal_static_cls"));
                                self.emit_u16(Op::LOCAL_SET, cls_tmp);
                                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                self.emit_u16(Op::LOCAL_GET, cls_tmp);
                                self.emit_u8(Op::CALL_REF, 1);
                                return Ok(());
                            }
                        }
                    }
                }

                // Dotted-name constants and the common namespace resolver, asked
                // without a language gate: `lookup_constant` reads the PROFILE's
                // own constant table and `resolve_profile_namespace_chain` the
                // profile's own namespace surface. A profile declaring neither
                // gets `None` from both and falls through unchanged.
                {
                    let parts = self.flatten_member_chain(expr);
                    if !parts.is_empty() {
                        let const_key = parts.join(".");
                        if let Some(cv) = self.profile.lookup_constant(&const_key).cloned() {
                            match cv {
                                ConstantValue::Bool(b) => self.emit_const(Value::Bool(b)),
                                ConstantValue::Float(f) => self.emit_const(Value::F64(f)),
                                ConstantValue::Str(s) => {
                                    self.emit_const(Value::String(Arc::from(s.as_str())))
                                }
                            }
                            return Ok(());
                        }
                        // namespaceplan.md: platform surfaces are data in the
                        // shared tree; the common resolver handles the mounted chain.
                        match self.resolve_profile_namespace_chain(&parts) {
                            Some(super::resolver::Resolution::GlobalAccess { name }) => {
                                self.emit_global_read(&name);
                                return Ok(());
                            }
                            Some(super::resolver::Resolution::Tree(
                                crate::primitives::namespaces::ResolutionTarget::CommonEmit(emit),
                            )) => {
                                self.emit_common(&emit, 0, self.line);
                                return Ok(());
                            }
                            Some(
                                super::resolver::Resolution::HostImport { module, func }
                                | super::resolver::Resolution::Tree(
                                    crate::primitives::namespaces::ResolutionTarget::HostCall {
                                        module,
                                        func,
                                        ..
                                    },
                                ),
                            ) => {
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, 0);
                                return Ok(());
                            }
                            Some(super::resolver::Resolution::Tree(
                                crate::primitives::namespaces::ResolutionTarget::Const(value),
                            )) => {
                                self.emit_const(value);
                                return Ok(());
                            }
                            Some(super::resolver::Resolution::ResolvedPrefix {
                                target,
                                suffix }) => {
                                match target {
                                    crate::primitives::namespaces::ResolutionTarget::CommonEmit(
                                        emit,
                                    ) => {
                                        self.emit_common(&emit, 0, self.line);
                                    }
                                    crate::primitives::namespaces::ResolutionTarget::HostCall {
                                        module,
                                        func,
                                        ..
                                    } => {
                                        let idx = self.import(&module, &func);
                                        self.emit_host_call(idx, 0);
                                    }
                                    crate::primitives::namespaces::ResolutionTarget::Const(
                                        value,
                                    ) => {
                                        self.emit_const(value);
                                    }
                                    _ => {}
                                }
                                for part in suffix {
                                    if matches!(part.as_str(), "length" | "count") {
                                        common::collections::emit_len(
                                            &mut self.chunks,
                                            self.current,
                                            self.line,
                                        );
                                        continue;
                                    }
                                    let idx = self.str_const(&part);
                                    self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                                }
                                return Ok(());
                            }
                            Some(super::resolver::Resolution::NamespaceChain {
                                parts: ns_parts }) => {
                                if ns_parts.len() >= 2 {
                                    let mut found_window: Option<(usize, usize)> = None;
                                    'outer: for start in 0..ns_parts.len().saturating_sub(1) {
                                        for end in ((start + 2)..=ns_parts.len()).rev() {
                                            let key = ns_parts[start..end].join(".");
                                            if self.profile.lookup_constant(&key).is_some() {
                                                found_window = Some((start, end));
                                                break 'outer;
                                            }
                                        }
                                    }
                                    if let Some((_const_start, const_end)) = found_window {
                                        let key = ns_parts[_const_start..const_end].join(".");
                                        let cv =
                                            self.profile.lookup_constant(&key).cloned().unwrap();
                                        match cv {
                                            ConstantValue::Bool(b) => {
                                                self.emit_const(Value::Bool(b))
                                            }
                                            ConstantValue::Float(f) => {
                                                self.emit_const(Value::F64(f))
                                            }
                                            ConstantValue::Str(s) => self
                                                .emit_const(Value::String(Arc::from(s.as_str()))) }
                                        for part in &ns_parts[const_end..] {
                                            let idx = self.str_const(part);
                                            self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                                        }
                                        return Ok(());
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }

                if self.class_prototype_dispatch()
                    && matches!(&object.kind, ExprKind::Super)
                    && !*null_safe
                {
                    if self.current_class.is_some() {
                        let result_slot = self.define_local("__js_super_prop_result");
                        let saved_this = self.save_js_this("__js_prev_this_super_prop");
                        self.emit_js_current_this_value();
                        self.set_js_this_from_stack();

                        let getter_key = self.str_const(&format!("__get_{}", field));
                        self.emit_js_super_home_base();
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, getter_key);
                        let getter_slot = self.define_local("__js_super_prop_getter");
                        self.emit_u16(Op::LOCAL_SET, getter_slot);
                        self.emit_u16(Op::LOCAL_GET, getter_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.emit_js_super_home_base();
                        let field_key = self.str_const(field);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, field_key);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, getter_slot);
                        self.emit_js_current_this_value();
                        self.emit_u8(Op::CALL_REF, 1);
                        self.chunk().emit_end(line);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.restore_js_this(saved_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        return Ok(());
                    }
                }

                // Proxy get-trap dispatch (JS profile, only when the
                // module references `Proxy` somewhere). Routes member
                // reads through the inline dispatcher in
                // `emitter/js/proxy_adapter.rs`. Non-proxy code keeps
                // the direct STRUCT_GET path.
                if self.uses_proxy && !*null_safe {
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    let line = self.line;
                    vybe_runtime::registry::hooks(&self.profile.name)
                        .proxy_get
                        .unwrap()(&mut self.chunks, self.current, line);
                    return Ok(());
                }

                if self.profile.supports_private_fields && field.starts_with('#') && !*null_safe {
                    self.compile_expr(object)?;
                    let obj_slot = self.define_local("__js_private_member_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    self.emit_const(Value::String(Arc::from(
                        "Cannot read private member from null or undefined",
                    )));
                    self.emit_js_exception_ctor_from_message_value("TypeError")?;
                    common::errors::emit_throw(self.chunk(), line);
                    self.chunk().emit_end(line);

                    let field_name = self.js_member_storage_name_for_receiver(object, field);
                    let getter_name = format!("__get_{}", field_name);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_const(Value::String(Arc::from(getter_name.as_str())));
                    // `has` (proto-walk, raw key) not `hasOwn`: the private
                    // accessor key is `__get_/__set___js_private_*` — a `__`
                    // key that `hasOwn` hides, and under prototype dispatch the
                    // accessor lives on the class prototype, not the instance.
                    let has_own_idx = self.import("ecma:object", "has");
                    self.emit_host_call(has_own_idx, 2);
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let getter_key = self.str_const(&getter_name);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, getter_key);
                    let getter_slot = self.define_local("__js_private_getter");
                    self.emit_u16(Op::LOCAL_SET, getter_slot);
                    let saved_this = self.save_js_this("__js_prev_this_private_get");
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, getter_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    let result_slot = self.define_local("__js_private_member_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.restore_js_this(saved_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);

                    self.chunk().emit_else(line);

                    self.emit_js_private_brand_check(obj_slot, &field_name)?;
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let prop = self.str_const(&field_name);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, prop);

                    self.chunk().emit_end(line);
                    return Ok(());
                }

                if self.profile.dynamic_member_access {
                    if *null_safe {
                        self.compile_expr(object)?;
                        let nullsafe_obj_slot = self.define_local("__js_member_nullsafe_obj");
                        self.emit_u16(Op::LOCAL_SET, nullsafe_obj_slot);
                        self.emit_u16(Op::LOCAL_GET, nullsafe_obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        common::expressions::emit_undefined(self.chunk(), line);
                        self.chunk().emit_else(line);
                        let obj_slot = nullsafe_obj_slot;
                        if field == "constructor" {
                            let constructor_of = self.import("ecma:value", "constructorOf");
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_host_call(constructor_of, 1);
                        } else if field == "length" {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            let prop = self.str_const("length");
                            self.emit_struct_field_op(Op::STRUCT_GET, 0, prop);
                        } else {
                            let field_name =
                                self.js_member_storage_name_for_receiver(object, field);
                            let prop = self.str_const(&field_name);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_struct_field_op(Op::STRUCT_GET, 0, prop);
                            let val_slot = self.define_local("__js_member_val");
                            self.emit_u16(Op::LOCAL_SET, val_slot);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            let lookup_line = self.line;
                            self.chunk().emit_if_value(lookup_line);
                            self.emit_member_get_from_value(obj_slot, &field_name);
                            self.chunk().emit_else(lookup_line);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            self.chunk().emit_end(lookup_line);
                        }
                        self.chunk().emit_end(line);
                    } else {
                        self.compile_expr(object)?;
                        let obj_slot = self.define_local("__js_member_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);

                        // Bind `__js_this = obj` so a getter installed
                        // by `Object.defineProperty` (which runs via
                        // STRUCT_GET's `__get_<name>` accessor dispatch)
                        // sees the receiver. The VM's accessor-call
                        // path doesn't set `__js_this` itself; doing it
                        // here keeps the semantics consistent with the
                        // explicit method-call path.
                        let saved_this = self.save_js_this("__js_prev_this_member");
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.set_js_this_from_stack();

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let msg_line = self.line;
                        self.chunk().emit_if_value(msg_line);
                        self.emit_const(Value::String(Arc::from(
                            "Cannot read properties of undefined",
                        )));
                        self.chunk().emit_else(msg_line);
                        self.emit_const(Value::String(Arc::from("Cannot read properties of null")));
                        self.chunk().emit_end(msg_line);
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        common::errors::emit_throw(self.chunk(), line);
                        self.chunk().emit_end(line);

                        let result_slot = self.define_local("__js_member_result");
                        if field == "constructor" {
                            let constructor_of = self.import("ecma:value", "constructorOf");
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_host_call(constructor_of, 1);
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.restore_js_this(saved_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            return Ok(());
                        }
                        if field == "length" {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            let prop = self.str_const("length");
                            self.emit_struct_field_op(Op::STRUCT_GET, 0, prop);
                            // §10.1.8.1 OrdinaryGet: a missing own
                            // `length` walks the prototype chain like any
                            // other key (e.g. AsyncFunction.prototype
                            // .length inherits %Function.prototype%'s 0).
                            let val_slot = self.define_local("__js_member_len");
                            self.emit_u16(Op::LOCAL_SET, val_slot);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            let lookup_line = self.line;
                            self.chunk().emit_if_value(lookup_line);
                            self.emit_member_get_from_value(obj_slot, "length");
                            self.chunk().emit_else(lookup_line);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            self.chunk().emit_end(lookup_line);
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.restore_js_this(saved_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            return Ok(());
                        }
                        let symbol_end = if field == "description" {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            fn_call!(self, "ecma:value", "typeof", 1);
                            self.emit_const(Value::String(Arc::from("symbol")));
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                            };
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            self.chunk().emit_if(line);
                            let invoke = self.import("ecma:value", "invokeMethod");
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_const(Value::String(Arc::from("description")));
                            self.emit_host_call(invoke, 2);
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.chunk().emit_else(line);
                            Some(line)
                        } else {
                            None
                        };

                        let field_name = self.js_member_storage_name_for_receiver(object, field);
                        let prop = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, prop);
                        let val_slot = self.define_local("__js_member_val");
                        self.emit_u16(Op::LOCAL_SET, val_slot);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let lookup_line = self.line;
                        self.chunk().emit_if_value(lookup_line);
                        self.emit_member_get_from_value(obj_slot, &field_name);
                        self.chunk().emit_else(lookup_line);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.chunk().emit_end(lookup_line);
                        // Restore the caller's __js_this — value already
                        // on stack as the access result.
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        if let Some(line) = symbol_end {
                            self.chunk().emit_end(line);
                        }
                        self.restore_js_this(saved_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    return Ok(());
                }

                // ONE receiver-typing path for every language. The two branches
                // were the same mechanism at different strengths — the gated
                // one is a strict superset (`resolve_receiver_type_hint` starts
                // from `lookup_var_type_hint` for an `Ident` and adds scope
                // types, global hints, static fields and type aliases, then
                // falls back to `infer_expr_type_hint` exactly as the other
                // branch did). Neither names a language; the flag only decided
                // who got the better answer.
                let receiver_type_hint =
                    crate::primitives::calls::resolve_receiver_type_hint(self, object)
                        .or_else(|| self.infer_expr_type_hint(object));
                if self.profile.member_invokes_parameterless_method && !*null_safe {
                    if let Some(class_name) = receiver_type_hint.as_deref().and_then(|type_hint| {
                        self.resolve_pending_class_name_for_type_hint(type_hint)
                    }) {
                        let member_name = self.canon(field);
                        let zero_arg_instance_method = self
                            .pending_classes
                            .get(&class_name)
                            .is_some_and(|pending| {
                                !pending.fields.iter().any(|name| name == &member_name)
                                    && pending
                                        .instance_method_overloads
                                        .get(&member_name)
                                        .is_some_and(|overloads| {
                                            overloads.iter().any(|overload| {
                                                overload.signature.min_arity == 0
                                                    && !overload.signature.has_rest
                                            })
                                        })
                            });
                        if zero_arg_instance_method {
                            self.compile_expr(object)?;
                            let obj_slot = self
                                .scope()
                                .resolve("__pascal_member_obj")
                                .unwrap_or_else(|| self.define_local("__pascal_member_obj"));
                            self.emit_u16(Op::LOCAL_SET, obj_slot);
                            let prop = self.str_const(&member_name);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_struct_field_op(Op::STRUCT_GET, 0, prop);
                            let fn_slot = self
                                .scope()
                                .resolve("__pascal_member_fn")
                                .unwrap_or_else(|| self.define_local("__pascal_member_fn"));
                            self.emit_u16(Op::LOCAL_SET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_u8(Op::CALL_REF, 1);
                            return Ok(());
                        }
                    }
                }

                let receiver_is_nullable = receiver_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim().ends_with('?'));

                // `Rank` plus an array-shaped hint IS the signal — the parse
                // below needs `[` … `]` in the type hint, which only an array
                // type produces. The family check added nothing the field name
                // and the hint did not already say.
                let receiver_array_rank = if field == "Rank" {
                    let inferred = receiver_type_hint.as_deref().and_then(|type_hint| {
                        let normalized = Self::normalize_type_hint(type_hint);
                        let start = normalized.find('[')?;
                        let rest = &normalized[start + 1..];
                        let end = rest.find(']')?;
                        Some(rest[..end].chars().filter(|ch| *ch == ',').count() + 1)
                    });
                    Some(inferred.unwrap_or(1))
                } else {
                    None
                };

                // `HasValue`/`Value` on a receiver whose DECLARED type is
                // nullable. `receiver_is_nullable` is read off the type hint the
                // shared AST carries, and both member names are matched
                // case-sensitively, so this is the nullable-wrapper protocol
                // rather than any one language's spelling of it.
                if receiver_is_nullable {
                    match field.as_str() {
                        "HasValue" => {
                            self.compile_expr(object)?;
                            self.emit(Op::REF_IS_NULL);
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                            };
                            return Ok(());
                        }
                        "Value" => {
                            self.compile_expr(object)?;
                            return Ok(());
                        }
                        _ => {}
                    }
                }

                if let Some(rank) = receiver_array_rank {
                    self.emit_const(Value::I32(rank as i32));
                    return Ok(());
                }

                // The receiver has a size member if the language's OWN
                // registered tree says the type declares one. That replaces the
                // family check; the spelling list below is still consulted for
                // the shapes no type node covers (arrays, strings).
                let receiver_is_collection_like = if field.eq_ignore_ascii_case("Length")
                    || field.eq_ignore_ascii_case("Count")
                {
                    let unknown_receiver_default = field.eq_ignore_ascii_case("Length");
                    let type_scopes = &self.profile.namespaces.type_scopes;
                    let is_collection_like_type = |type_hint: &str| {
                        // The tree is the authority: a registered type that
                        // declares `Count` IS a collection, whichever platform
                        // registered it. This is what retires the hardcoded
                        // .NET name list that used to live here — a
                        // per-language table inside a twelve-language crate.
                        if vybe_runtime::namespaces::lookup_type_instance_member(
                            type_scopes,
                            type_hint,
                            "Count",
                        )
                        .is_some()
                        {
                            return true;
                        }
                        let normalized = Self::normalize_type_hint(type_hint);
                        let generic_start = normalized
                            .find('<')
                            .into_iter()
                            .chain(normalized.find("(of"))
                            .min()
                            .unwrap_or(normalized.len());
                        let bare = normalized[..generic_start].trim_end_matches('?').trim();
                        let terminal = bare.rsplit('.').next().unwrap_or(bare);
                        Self::is_string_type_hint(type_hint)
                            || matches!(
                                terminal,
                                "list"
                                    | "arraylist"
                                    | "dictionary"
                                    | "sorteddictionary"
                                    | "queue"
                                    | "stack"
                                    | "hashset"
                                    | "sortedset"
                                    | "set"
                                    | "collection"
                                    | "observablecollection"
                                    | "readonlyobservablecollection"
                                    | "icollection"
                                    | "readonlycollection"
                                    | "enumerable"
                                    | "ienumerable"
                                    | "readonlylist"
                                    | "ilist"
                                    | "array"
                            )
                            || bare.ends_with("[]")
                    };

                    match &object.kind {
                        ExprKind::Ident(_)
                        | ExprKind::New { .. }
                        | ExprKind::Call { .. }
                        | ExprKind::Cast { .. }
                        | ExprKind::Index { .. } => receiver_type_hint
                            .as_deref()
                            .map(is_collection_like_type)
                            .unwrap_or(unknown_receiver_default),
                        ExprKind::Lit(Literal::Str(_))
                        | ExprKind::Interpolation(_)
                        | ExprKind::Array(_) => true,
                        _ => unknown_receiver_default }
                } else {
                    false
                };

                let is_dotnet_dictionary_accessor = !self.case_sensitive
                    && matches!(field.as_str(), "Keys" | "Values")
                    && receiver_type_hint
                        .as_deref()
                        .map(|type_hint| {
                            Self::is_dictionary_type_hint(type_hint)
                                || Self::is_sorted_dictionary_type_hint(type_hint)
                        })
                        .unwrap_or(false)
                    && !matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if name.chars().next().map_or(false, |c| c.is_ascii_uppercase())
                    );

                if is_dotnet_dictionary_accessor {
                    self.compile_expr(object)?;
                    if field == "Keys" {
                        common::collections::emit_iter_keys(
                            &mut self.chunks,
                            self.current,
                            self.line,
                        );
                    } else {
                        self.emit_common("dict.values", 1, self.line);
                    }
                    return Ok(());
                }

                // Gated by the RECEIVER being collection-like, which the tree
                // now answers — not by which family compiled the file.
                let is_csharp_len_accessor = (field.eq_ignore_ascii_case("Length")
                        || field.eq_ignore_ascii_case("Count"))
                    && receiver_is_collection_like
                    && !matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if name.chars().next().map_or(false, |c| c.is_ascii_uppercase())
                    );
                let is_csharp_runtime_count_accessor = field == "Count"
                    && !is_csharp_len_accessor
                    && !*null_safe
                    && !matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                        if name.chars().next().map_or(false, |c| c.is_ascii_uppercase())
                    );

                // The type hint names the set types exactly; the family check
                // decided nothing the hint did not.
                if field == "Count"
                    && receiver_type_hint
                        .as_deref()
                        .map(|type_hint| {
                            let normalized = Self::normalize_type_hint(type_hint);
                            normalized.contains("hashset") || normalized.contains("sortedset")
                        })
                        .unwrap_or(false)
                    && !matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if name.chars().next().map_or(false, |c| c.is_ascii_uppercase())
                    )
                {
                    self.compile_expr(object)?;
                    fn_call!(self, "ecma:set", "size", 1);
                    return Ok(());
                }

                if !*null_safe && matches!(&object.kind, ExprKind::Super) {
                    if let Some(field_name) = self.field_storage_name_for_receiver(object, field) {
                        if self.emit_self_ref() {
                            let idx = self.str_const(&field_name);
                            self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                            return Ok(());
                        }
                    }
                }

                // The type-hint test below already names the type exactly, and
                // resolves it through the tree — nothing left for a family
                // check to decide.
                let is_dotnet_observable_count = field == "Count"
                    && receiver_type_hint
                        .as_deref()
                        .map(|type_hint| {
                            let normalized = Self::normalize_type_hint(type_hint);
                            normalized.contains("observablecollection")
                                || self
                                    .namespace_tree_instance_method_owner(type_hint, "Count", 0)
                                    .is_some_and(|owner| {
                                        owner.to_ascii_lowercase().contains("observablecollection")
                                    })
                        })
                        .unwrap_or(false);

                if is_csharp_len_accessor {
                    self.compile_expr(object)?;
                    if *null_safe {
                        inst!(self, core_wasm::dup);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.emit(Op::DROP);
                        self.emit_null();
                        self.chunk().emit_else(line);
                        common::collections::emit_len(&mut self.chunks, self.current, self.line);
                        self.chunk().emit_end(line);
                    } else {
                        common::collections::emit_len(&mut self.chunks, self.current, self.line);
                    }
                    return Ok(());
                } else if is_dotnet_observable_count {
                    self.compile_expr(object)?;
                    self.emit_common("dotnet.observable_collection_count", 1, self.line);
                    return Ok(());
                } else if field == "Items"
                    && receiver_type_hint
                        .as_deref()
                        .map(|type_hint| {
                            let normalized = Self::normalize_type_hint(type_hint);
                            normalized.contains("observablecollection")
                                || self
                                    .namespace_tree_instance_method_owner(type_hint, "Items", 0)
                                    .is_some_and(|owner| {
                                        owner.to_ascii_lowercase().contains("observablecollection")
                                    })
                        })
                        .unwrap_or(false)
                {
                    self.compile_expr(object)?;
                    self.emit_common("dotnet.observable_collection_items", 1, self.line);
                    return Ok(());
                } else if !self.profile.namespaces.type_scopes.is_empty()
                    && !*null_safe
                    && let Some(target) = receiver_type_hint.as_deref().and_then(|type_hint| {
                        let class_name = Self::normalize_type_hint(type_hint);
                        vybe_runtime::namespaces::lookup_type_property_target(
                            &self.profile.namespaces.type_scopes,
                            &class_name,
                            field,
                        )
                    })
                {
                    self.compile_expr(object)?;
                    match target {
                        vybe_runtime::component_model::InstancePropertyTarget::Host {
                            module,
                            func,
                            key } => {
                            let idx = self.import(&module, &func);
                            if let Some(key) = key {
                                self.emit_const(Value::String(Arc::from(key.as_str())));
                                self.emit_host_call(idx, 2);
                            } else {
                                self.emit_host_call(idx, 1);
                            }
                            return Ok(());
                        }
                        vybe_runtime::component_model::InstancePropertyTarget::Common { emit } => {
                            self.emit_common(&emit, 1, self.line);
                            return Ok(());
                        }
                    }
                } else {
                    self.compile_expr(object)?;
                    if !Self::is_pointer_runtime_field(field) {
                        self.emit_autoderef_pointer_cell();
                    }
                }

                let namespace_tree_zero_arg_method = if !self.profile.namespaces.type_scopes.is_empty()
                    && !*null_safe
                    && !is_csharp_len_accessor
                    && !is_csharp_runtime_count_accessor
                {
                    receiver_type_hint.as_deref().and_then(|type_hint| {
                        let class_name = Self::normalize_type_hint(type_hint);
                        vybe_runtime::namespaces::lookup_type_instance_target(
                            &self.profile.namespaces.type_scopes,
                            &class_name,
                            field,
                            0,
                        )
                    })
                } else {
                    None
                };

                if let Some(target) = namespace_tree_zero_arg_method {
                    if let vybe_runtime::component_model::InstanceMethodTarget::Common {
                        emit,
                        ..
                    } = &target
                    {
                        let line = self.line;
                        self.compile_expr(object)?;
                        self.emit_common(emit, 1, line);
                        return Ok(());
                    }

                    let obj_slot = self.define_local("__dotnet_zero_arg_obj");
                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    let value_slot = self.define_local("__dotnet_zero_arg_value");
                    let field_name = self.canon(field);
                    let canonical_idx = self.str_const(&field_name);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, canonical_idx);
                    self.emit_u16(Op::LOCAL_SET, value_slot);

                    // `field_name` is the name resolution settled on; a
                    // difference from the source spelling is the signal, and it
                    // is data, not a language family.
                    if field.as_str() != field_name {
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if(line);

                        let exact_idx = self.str_const(field);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, exact_idx);
                        self.emit_u16(Op::LOCAL_SET, value_slot);

                        self.chunk().emit_end(line);
                    }

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    let line = self.line;
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_else(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    match target {
                        vybe_runtime::component_model::InstanceMethodTarget::Host {
                            module,
                            func,
                            ..
                        } => {
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, 1);
                        }
                        vybe_runtime::component_model::InstanceMethodTarget::Common {
                            emit,
                            ..
                        } => {
                            let line = self.line;
                            self.emit_common(&emit, 1, line);
                        }
                    }
                    self.chunk().emit_end(line);
                    return Ok(());
                }

                if is_csharp_len_accessor {
                    common::collections::emit_len(&mut self.chunks, self.current, self.line);
                    return Ok(());
                }

                if is_csharp_runtime_count_accessor {
                    let obj_slot = self.define_local("__count_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                    let value_slot = self.define_local("__count_value");
                    self.emit_u16(Op::LOCAL_SET, value_slot);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    fn_call!(self, "ecma:value", "typeof", 1);
                    self.emit_const(Value::String(Arc::from("function")));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_end(line);
                    return Ok(());
                }

                // `reflection_type_metadata` is a registry lookup that names no
                // language and answers `None` for a type nothing registered, so
                // the gate only decided who was allowed to ask.
                let static_field_owner = {
                    receiver_type_hint.as_deref().and_then(|type_hint| {
                        let trimmed_type_hint = type_hint.trim().trim_end_matches("()").trim();
                        let metadata_type_hint = self
                            .reflection_type_metadata(type_hint)
                            .map(|_| type_hint)
                            .or_else(|| {
                                self.reflection_type_metadata(trimmed_type_hint)
                                    .map(|_| trimmed_type_hint)
                            })?;
                        self.reflection_type_metadata(metadata_type_hint)
                            .and_then(|meta| {
                                meta.fields.iter().find(|(name, field_meta)| {
                                    name.eq_ignore_ascii_case(field) && field_meta.is_static
                                })
                            })
                            .map(|_| {
                                let short_name =
                                    self.reflection_type_short_name(metadata_type_hint);
                                if self.defined_globals.contains(&self.canon(&short_name)) {
                                    short_name
                                } else {
                                    self.reflection_type_lookup_name(metadata_type_hint)
                                }
                            })
                    })
                };

                // Late-bound member read: try the instance field, fall back to
                // the type's static field. Triggered purely by having resolved a
                // static-field owner (data), not by language name — a member
                // whose static owner is known but instance value may be absent.
                if let Some(type_name) = static_field_owner {
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    let obj_slot = self.define_local("__member_static_fallback_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                    let value_slot = self.define_local("__member_static_fallback_value");
                    self.emit_u16(Op::LOCAL_SET, value_slot);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    let line = self.line;
                    self.chunk().emit_if_value(line);

                    self.emit_global_read(&self.canon(&type_name));
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.chunk().emit_end(line);
                    return Ok(());
                }

                if *null_safe && self.profile.namespaces.use_dotnet && !is_csharp_len_accessor {
                    let obj_slot = self.define_local("__dotnet_nullsafe_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.emit(Op::I32_EQZ);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let field_name = self
                        .field_storage_name_for_receiver(object, field)
                        .unwrap_or_else(|| self.canon(field));
                    let idx = self.str_const(&field_name);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                    self.chunk().emit_else(line);
                    self.emit_null();
                    self.chunk().emit_end(line);
                } else if *null_safe {
                    let obj_slot = self.define_local("__member_nullsafe_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    inst!(self, recipes::is_object);
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                    self.chunk().emit_else(line);
                    self.emit_null();
                    self.chunk().emit_end(line);
                } else {
                    let field_name = self
                        .field_storage_name_for_receiver(object, field)
                        .unwrap_or_else(|| self.canon(field));
                    // `field_name` is the name resolution settled on; a
                    // difference from the source spelling is the signal, and it
                    // is data, not a language family.
                    if field.as_str() != field_name {
                        let obj_slot = self.define_local("__dotnet_member_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);

                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                        let value_slot = self.define_local("__dotnet_member_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot);

                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if_value(line);

                        let exact_idx = self.str_const(field);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, exact_idx);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.chunk().emit_end(line);
                    } else {
                        let idx = self.str_const(&field_name);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                    }
                }
            }

            // ── Index access ────────────────────────────────────────────
            ExprKind::Index {
                object,
                index,
                null_safe } => {
                // `_G[k]` / `$GLOBALS[$k]` / `globals()[k]` — the module's
                // global namespace indexed by a RUNTIME key. Handled before the
                // ordinary index paths because the namespace has no object to
                // index: `_G` on its own evaluates to nil today.
                if self.expr_is_global_namespace(object) {
                    return self.emit_global_namespace_index_get(index);
                }
                // A class declaring `operator []` / `__getitem__` makes `x[i]`
                // a method call on that type. The receiver's declared type
                // settles it at compile time, so this emits a direct call and
                // every other receiver — array, dict, string — keeps the plain
                // index path below with no added runtime check.
                if !*null_safe
                    && !matches!(index.kind, ExprKind::Range { .. } | ExprKind::Slice { .. })
                    && self.expr_has_user_indexer(object)
                {
                    let line = self.line;
                    let obj_slot = self.define_local("__idx_recv");
                    let key_slot = self.define_local("__idx_key");
                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.compile_expr(index)?;
                    self.emit_u16(Op::LOCAL_SET, key_slot);
                    let getter = self.str_const(&vybe_ast::protocol_slot_key(
                        vybe_ast::ProtocolSlot::GetItem,
                    ));
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, getter);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    self.chunk().emit_op_u8(Op::CALL_REF, 2, line);
                    return Ok(());
                }
                // A Range used as the index is a slice operation
                // (C# `arr[1..3]` / `s[0..5]`, Python `arr[1:3]` / `s[0:5]`).
                // Route through compiler_common's polymorphic slice helper so
                // strings and arrays both work uniformly.
                if let ExprKind::Range {
                    start,
                    end,
                    inclusive } = &index.kind
                {
                    let line = self.line;
                    if self.profile.ecma_array_method_dispatch {
                        // Emit an inline polymorphic slice for JS: strings use
                        // wasm:js-string.substring, arrays call ecma:array.slice.
                        // Save operands to locals so we can test the receiver.
                        let obj_slot = self.define_local("__js_range_slice_obj");
                        let start_slot = self.define_local("__js_range_slice_start");
                        let end_slot = self.define_local("__js_range_slice_end");

                        self.compile_expr(object)?;
                        self.emit_u16(Op::LOCAL_SET, obj_slot);

                        self.compile_expr(start)?;
                        self.emit_u16(Op::LOCAL_SET, start_slot);

                        self.compile_expr(end)?;
                        // Inclusive `a..b` slice → exclusive upper bound b+1.
                        if *inclusive {
                            inst!(self, core_wasm::i32_const, 1);
                            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                        }
                        self.emit_u16(Op::LOCAL_SET, end_slot);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        fn_call!(self, "wasm:js-string", "test", 1);
                        self.chunk().emit_if_value(line);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.emit_u16(Op::LOCAL_GET, end_slot);
                        fn_call!(self, "wasm:js-string", "substring", 3);

                        self.chunk().emit_else(line);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.emit_u16(Op::LOCAL_GET, end_slot);
                        common::collections::emit_slice(&mut self.chunks, self.current, line);

                        self.chunk().emit_end(line);
                    } else {
                        let line = self.line;
                        self.compile_expr(object)?;
                        self.compile_expr(start)?;
                        self.compile_expr(end)?;
                        // Inclusive `a..b` slice → exclusive upper bound b+1.
                        if *inclusive {
                            inst!(self, core_wasm::i32_const, 1);
                            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                        }
                        // Direct polymorphic `ecma:array.slice` (string→substring,
                        // array→slice) — no `__vybe_slice` func-ref/chunk detour.
                        common::collections::emit_slice(&mut self.chunks, self.current, line);
                    }
                } else if let ExprKind::Slice { lower, upper, step } = &index.kind {
                    self.compile_expr(object)?;
                    let line = self.line;
                    if step.is_none() {
                        if self.profile.ecma_array_method_dispatch {
                            // JS: compute start/end into locals, then dispatch
                            // to wasm:js-string.substring for strings or
                            // ecma:array.slice for arrays.
                            let obj_slot = self.define_local("__js_index_slice_obj");
                            let start_slot = self.define_local("__js_index_slice_start");
                            let end_slot = self.define_local("__js_index_slice_end");

                            self.emit_u16(Op::LOCAL_SET, obj_slot);

                            if let Some(l) = lower {
                                self.compile_expr(l)?;
                            } else {
                                inst!(self, core_wasm::i32_const, 0);
                            }
                            self.emit_u16(Op::LOCAL_SET, start_slot);

                            if let Some(u) = upper {
                                self.compile_expr(u)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_slot);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            self.emit_u16(Op::LOCAL_SET, end_slot);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            fn_call!(self, "wasm:js-string", "test", 1);
                            self.chunk().emit_if_value(line);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_u16(Op::LOCAL_GET, start_slot);
                            self.emit_u16(Op::LOCAL_GET, end_slot);
                            fn_call!(self, "wasm:js-string", "substring", 3);

                            self.chunk().emit_else(line);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_u16(Op::LOCAL_GET, start_slot);
                            self.emit_u16(Op::LOCAL_GET, end_slot);
                            common::collections::emit_slice(&mut self.chunks, self.current, line);

                            self.chunk().emit_end(line);
                        } else if self.profile.slice_bounds_inclusive {
                            let obj_slot = self.define_local("__fortran_index_slice_obj");
                            self.emit_u16(Op::LOCAL_SET, obj_slot);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            if let Some(l) = lower {
                                self.compile_expr(l)?;
                            } else {
                                inst!(self, core_wasm::i32_const, 0);
                            }
                            if let Some(u) = upper {
                                self.compile_expr(u)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_slot);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            crate::primitives::slices::emit_contiguous(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                        } else {
                            let obj_slot = self.define_local("__py_index_slice_obj");
                            self.emit_u16(Op::LOCAL_SET, obj_slot);

                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            if let Some(l) = lower {
                                self.compile_expr(l)?;
                            } else {
                                inst!(self, core_wasm::i32_const, 0);
                            }
                            if let Some(u) = upper {
                                self.compile_expr(u)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_slot);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            crate::primitives::slices::emit_contiguous(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                        }
                    } else {
                        let step_const = step.as_ref().and_then(|expr| match &expr.kind {
                            ExprKind::Lit(Literal::Int(n)) => Some(*n),
                            ExprKind::Unary {
                                op: UnaryOp::Neg,
                                expr } => match &expr.kind {
                                ExprKind::Lit(Literal::Int(n)) => Some(-*n),
                                _ => None },
                            _ => None });

                        if lower.is_none() && upper.is_none() {
                            if step_const == Some(-1) {
                                inst!(self, core_wasm::dup);
                                fn_call!(self, "wasm:js-string", "test", 1);
                                let line = self.line;
                                self.chunk().emit_if_value(line);
                                {
                                    let l = self.line;
                                    crate::primitives::strings::emit_str_reverse(self.chunk(), l);
                                }
                                self.chunk().emit_else(line);
                                self.emit_null();
                                self.emit_null();
                                if let Some(s) = step {
                                    self.compile_expr(s)?;
                                } else {
                                    self.emit_null();
                                }
                                {
                                    let opts = crate::primitives::slices::Options::new(
                                        self.profile.slice_step_zero_raises,
                                    );
                                    crate::primitives::slices::emit_stepped(
                                        &mut self.chunks,
                                        self.current,
                                        line,
                                        opts,
                                    );
                                }
                                self.chunk().emit_end(line);
                                return Ok(());
                            }

                            if let Some(step_value) = step_const.filter(|n| *n > 1) {
                                inst!(self, core_wasm::dup);
                                fn_call!(self, "wasm:js-string", "test", 1);
                                let line = self.line;
                                self.chunk().emit_if_value(line);

                                let str_slot = self.define_local("__py_stride_string");
                                let result_slot = self.define_local("__py_stride_result");
                                let index_slot = self.define_local("__py_stride_index");
                                let len_slot = self.define_local("__py_stride_len");

                                self.emit_u16(Op::LOCAL_SET, str_slot);
                                self.emit_const(Value::String(Arc::from("")));
                                self.emit_u16(Op::LOCAL_SET, result_slot);
                                inst!(self, core_wasm::i32_const, 0);
                                self.emit_u16(Op::LOCAL_SET, index_slot);
                                self.emit_u16(Op::LOCAL_GET, str_slot);
                                fn_call!(self, "wasm:js-string", "length", 1);
                                self.emit_u16(Op::LOCAL_SET, len_slot);

                                let stride_block = self.chunk().emit_block(line);
                                let (stride_loop, _) = self.chunk().emit_loop_s(line);
                                self.emit_u16(Op::LOCAL_GET, index_slot);
                                self.emit_u16(Op::LOCAL_GET, len_slot);
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                                };
                                {
                                    let line = self.line;
                                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                                };
                                self.chunk().emit_br_if(1, line);

                                self.emit_u16(Op::LOCAL_GET, result_slot);
                                self.emit_u16(Op::LOCAL_GET, str_slot);
                                self.emit_u16(Op::LOCAL_GET, index_slot);
                                self.emit(Op::F64_FROM_I32);
                                fn_call!(self, "ecma:string", "charAt", 2);
                                common::strings::emit_str_concat(self.chunk(), line);
                                self.emit_u16(Op::LOCAL_SET, result_slot);

                                self.emit_u16(Op::LOCAL_GET, index_slot);
                                self.emit_const(Value::I32(step_value as i32));
                                self.emit(Op::I32_ADD);
                                self.emit_u16(Op::LOCAL_SET, index_slot);
                                self.chunk().emit_br(0, line);
                                self.chunk().emit_end(line);
                                self.chunk().patch_loop(stride_loop);
                                self.chunk().emit_end(line);
                                self.chunk().patch_block(stride_block);
                                self.emit_u16(Op::LOCAL_GET, result_slot);

                                self.chunk().emit_else(line);
                                self.emit_null();
                                self.emit_null();
                                if let Some(s) = step {
                                    self.compile_expr(s)?;
                                } else {
                                    self.emit_null();
                                }
                                {
                                    let opts = crate::primitives::slices::Options::new(
                                        self.profile.slice_step_zero_raises,
                                    );
                                    crate::primitives::slices::emit_stepped(
                                        &mut self.chunks,
                                        self.current,
                                        line,
                                        opts,
                                    );
                                }
                                self.chunk().emit_end(line);
                                return Ok(());
                            }
                        }

                        if let Some(l) = lower {
                            self.compile_expr(l)?;
                        } else {
                            self.emit_null();
                        }
                        if let Some(u) = upper {
                            self.compile_expr(u)?;
                        } else {
                            self.emit_null();
                        }
                        if let Some(s) = step {
                            self.compile_expr(s)?;
                        } else {
                            self.emit_null();
                        }
                        {
                            let opts = crate::primitives::slices::Options::new(
                                self.profile.slice_step_zero_raises,
                            );
                            crate::primitives::slices::emit_stepped(
                                &mut self.chunks,
                                self.current,
                                line,
                                opts,
                            );
                        }
                    }
                } else if self.profile.string_index_is_one_based
                    && self.expr_is_known_string_receiver(object)
                {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    self.emit_const(Value::F64(1.0));
                    self.emit(Op::F64_SUB);
                    fn_call!(self, "ecma:string", "charAt", 2);
                } else if self.uses_proxy && !*null_safe {
                    // Proxy get-trap dispatch on bracket-notation reads.
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    let line = self.line;
                    vybe_runtime::registry::hooks(&self.profile.name)
                        .proxy_get
                        .unwrap()(&mut self.chunks, self.current, line);
                } else if self.profile.dynamic_member_access && *null_safe {
                    self.compile_expr(object)?;
                    let obj_slot = self.define_local("__js_index_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    common::expressions::emit_undefined(self.chunk(), line);
                    self.chunk().emit_else(line);
                    let key_slot = self.define_local("__js_index_key");
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    match &index.kind {
                        ExprKind::Member {
                            object,
                            field,
                            null_safe: false } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") => {
                            let fallback_key = match field.as_str() {
                                "iterator" => Some("iterator"),
                                "asyncIterator" => Some("asyncIterator"),
                                "toPrimitive" => Some("toprimitive"),
                                "hasInstance" => Some("hasinstance"),
                                _ => None };
                            if let Some(fallback_key) = fallback_key {
                                self.emit_const(Value::String(Arc::from(fallback_key)));
                            } else {
                                self.compile_expr(index)?;
                            }
                        }
                        _ => self.compile_expr(index)? }
                    if self.profile.negative_index_wraps {
                        self.emit_negative_index_wrap();
                    }
                    self.emit_u16(Op::LOCAL_SET, key_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    if let Some(ns) = self
                        .infer_expr_type_hint(object)
                        .as_deref()
                        .map(Self::normalize_type_hint)
                        .and_then(|type_hint| match type_hint.as_str() {
                            "bigint64array" => Some("ecma:bigint64array"),
                            "biguint64array" => Some("ecma:biguint64array"),
                            _ => None })
                    {
                        let idx = self.import(ns, "get");
                        self.emit_host_call(idx, 2);
                    } else {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    let val_slot = self.define_local("__js_index_val");
                    self.emit_u16(Op::LOCAL_SET, val_slot);
                    // String[n] out-of-bounds: ARRAY_GET returns null, but JS spec (§6.1.4.1) needs undefined.
                    {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        fn_call!(self, "wasm:js-string", "test", 1);
                        let string_line = self.line;
                        self.chunk().emit_if(string_line);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.emit(Op::REF_IS_NULL);
                        let null_line = self.line;
                        self.chunk().emit_if(null_line);
                        inst!(self, core_wasm::undefined);
                        self.emit_u16(Op::LOCAL_SET, val_slot);
                        self.chunk().emit_end(null_line);
                        self.chunk().emit_end(string_line);
                    }
                    self.emit_u16(Op::LOCAL_GET, val_slot);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    let lookup_line = self.line;
                    self.chunk().emit_if_value(lookup_line);
                    // STRUCT_GET missed → direct ecma `[[Get]]` (Reflect.get),
                    // guarding non-object receivers → undefined (Reflect.get
                    // throws on non-object; __vybe_js_get_method returned undefined).
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    inst!(self, recipes::is_object);
                    self.chunk().emit_if_value(lookup_line);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    match &index.kind {
                        ExprKind::Member {
                            object,
                            field,
                            null_safe: false } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") => {
                            let fallback_key = match field.as_str() {
                                "iterator" => Some("iterator"),
                                "asyncIterator" => Some("asyncIterator"),
                                "toPrimitive" => Some("toprimitive"),
                                "hasInstance" => Some("hasinstance"),
                                _ => None };
                            if let Some(fallback_key) = fallback_key {
                                self.emit_const(Value::String(Arc::from(fallback_key)));
                            } else {
                                self.emit_u16(Op::LOCAL_GET, key_slot);
                            }
                        }
                        _ => self.emit_u16(Op::LOCAL_GET, key_slot) }
                    let reflect_idx = self.import("ecma:reflect", "get");
                    self.emit_host_call(reflect_idx, 2);
                    self.chunk().emit_else(lookup_line);
                    inst!(self, core_wasm::undefined);
                    self.chunk().emit_end(lookup_line);
                    self.chunk().emit_else(lookup_line);
                    self.emit_u16(Op::LOCAL_GET, val_slot);
                    self.chunk().emit_end(lookup_line);
                    self.chunk().emit_end(line);
                } else if self.profile.namespaces.use_dotnet {
                    if self
                            .infer_expr_type_hint(object)
                            .as_deref()
                            .map(Self::normalize_type_hint)
                            .is_some_and(|type_hint| {
                                type_hint
                                    .rsplit('.')
                                    .next()
                                    .is_some_and(|name| name.eq_ignore_ascii_case("StringBuilder"))
                            })
                    {
                        self.compile_expr(object)?;
                        self.compile_collection_key(object, index)?;
                        let line = self.line;
                        self.emit_common("dotnet.sb_index_get", 2, line);
                        return Ok(());
                    }

                    self.compile_expr(object)?;
                    let obj_slot = self.define_local("__index_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    let null_safe_if = if *null_safe {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.emit_null();
                        self.chunk().emit_else(line);
                        Some(line)
                    } else {
                        None
                    };

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let getter_key = self.str_const("__get___index__");
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, getter_key);
                    let getter_slot = self.define_local("__index_getter");
                    self.emit_u16(Op::LOCAL_SET, getter_slot);

                    self.emit_u16(Op::LOCAL_GET, getter_slot);
                    self.emit(Op::REF_IS_NULL);
                    let getter_line = self.line;
                    self.chunk().emit_if_value(getter_line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.compile_collection_key(object, index)?;
                    if self.profile.negative_index_wraps {
                        self.emit_negative_index_wrap();
                    }
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    self.chunk().emit_else(getter_line);
                    self.emit_u16(Op::LOCAL_GET, getter_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.compile_collection_key(object, index)?;
                    self.emit_u8(Op::CALL_REF, 2);
                    self.chunk().emit_end(getter_line);
                    if let Some(line) = null_safe_if {
                        self.chunk().emit_end(line);
                    }
                } else {
                    if self.profile.dynamic_member_access {
                        self.compile_expr(object)?;
                        let obj_slot = self.define_local("__js_index_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let msg_line = self.line;
                        self.chunk().emit_if_value(msg_line);
                        self.emit_const(Value::String(Arc::from(
                            "Cannot read properties of undefined",
                        )));
                        self.chunk().emit_else(msg_line);
                        self.emit_const(Value::String(Arc::from("Cannot read properties of null")));
                        self.chunk().emit_end(msg_line);
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        common::errors::emit_throw(self.chunk(), line);
                        self.chunk().emit_end(line);

                        let key_slot = self.define_local("__js_index_key");
                        match &index.kind {
                            ExprKind::Member {
                                object,
                                field,
                                null_safe: false } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") =>
                            {
                                let fallback_key = match field.as_str() {
                                    "iterator" => Some("iterator"),
                                    "asyncIterator" => Some("asyncIterator"),
                                    "toPrimitive" => Some("toprimitive"),
                                    "hasInstance" => Some("hasinstance"),
                                    _ => None };
                                if let Some(fallback_key) = fallback_key {
                                    self.emit_const(Value::String(Arc::from(fallback_key)));
                                } else {
                                    self.compile_array_index_operand_for_owner(object, index)?;
                                }
                            }
                            _ => self.compile_array_index_operand_for_owner(object, index)? }
                        if self.profile.negative_index_wraps {
                            self.emit_negative_index_wrap();
                        }
                        self.emit_u16(Op::LOCAL_SET, key_slot);

                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        self.emit_u16(Op::LOCAL_GET, key_slot);
                        if let Some(ns) = self
                            .infer_expr_type_hint(object)
                            .as_deref()
                            .map(Self::normalize_type_hint)
                            .and_then(|type_hint| match type_hint.as_str() {
                                "bigint64array" => Some("ecma:bigint64array"),
                                "biguint64array" => Some("ecma:biguint64array"),
                                _ => None })
                        {
                            let idx = self.import(ns, "get");
                            self.emit_host_call(idx, 2);
                        } else {
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        let val_slot = self.define_local("__js_index_val");
                        self.emit_u16(Op::LOCAL_SET, val_slot);
                        // String[n] out-of-bounds: ARRAY_GET returns null, but JS spec needs undefined.
                        {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            fn_call!(self, "wasm:js-string", "test", 1);
                            let string_line = self.line;
                            self.chunk().emit_if(string_line);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            self.emit(Op::REF_IS_NULL);
                            let null_line = self.line;
                            self.chunk().emit_if(null_line);
                            inst!(self, core_wasm::undefined);
                            self.emit_u16(Op::LOCAL_SET, val_slot);
                            self.chunk().emit_end(null_line);
                            self.chunk().emit_end(string_line);
                        }
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let lookup_line = self.line;
                        self.chunk().emit_if_value(lookup_line);
                        // STRUCT_GET missed → direct ecma `[[Get]]` (Reflect.get),
                        // guarding non-object receivers → undefined.
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        inst!(self, recipes::is_object);
                        self.chunk().emit_if_value(lookup_line);
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        match &index.kind {
                            ExprKind::Member {
                                object,
                                field,
                                null_safe: false } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") =>
                            {
                                let fallback_key = match field.as_str() {
                                    "iterator" => Some("iterator"),
                                    "asyncIterator" => Some("asyncIterator"),
                                    "toPrimitive" => Some("toprimitive"),
                                    "hasInstance" => Some("hasinstance"),
                                    _ => None };
                                if let Some(fallback_key) = fallback_key {
                                    self.emit_const(Value::String(Arc::from(fallback_key)));
                                } else {
                                    self.emit_u16(Op::LOCAL_GET, key_slot);
                                }
                            }
                            _ => self.emit_u16(Op::LOCAL_GET, key_slot) }
                        let reflect_idx = self.import("ecma:reflect", "get");
                        self.emit_host_call(reflect_idx, 2);
                        self.chunk().emit_else(lookup_line);
                        inst!(self, core_wasm::undefined);
                        self.chunk().emit_end(lookup_line);
                        self.chunk().emit_else(lookup_line);
                        self.emit_u16(Op::LOCAL_GET, val_slot);
                        self.chunk().emit_end(lookup_line);
                        return Ok(());
                    }
                    // Redundant name check removed: `x.__base[x.__idx]` with the
                    // same owner is a shape only the C walker builds.
                    let is_c_pointer_base_index = {
                        match (&object.kind, &index.kind) {
                            (
                                ExprKind::Member {
                                    object: base_owner,
                                    field: base_field,
                                    ..
                                },
                                ExprKind::Member {
                                    object: idx_owner,
                                    field: idx_field,
                                    ..
                                },
                            ) => {
                                matches!(
                                    (&base_owner.kind, &idx_owner.kind, base_field.as_str(), idx_field.as_str()),
                                    (ExprKind::Ident(a), ExprKind::Ident(b), "__base", "__idx") if a == b
                                )
                            }
                            _ => false }
                    };

                    if is_c_pointer_base_index {
                        self.compile_expr(object)?;
                        let obj_tmp = self.define_local("__pointer_index_get_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);

                        self.compile_array_index_operand_for_owner(object, index)?;
                        if self.profile.negative_index_wraps {
                            self.emit_negative_index_wrap();
                        }
                        let key_tmp = self.define_local("__pointer_index_get_key");
                        self.emit_u16(Op::LOCAL_SET, key_tmp);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        inst!(self, recipes::is_object);
                        let line = self.line;
                        self.chunk().emit_if(line);

                        let kind_key = self.str_const("__ref_kind");
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, kind_key);
                        self.emit_string_eq_literal("cell");
                        let line = self.line;
                        self.chunk().emit_if(line);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        crate::primitives::references::emit_cell_load(
                            &mut self.chunks,
                            self.current,
                            self.line,
                        );

                        let line = self.line;
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.chunk().emit_end(line);

                        let line = self.line;
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.chunk().emit_end(line);
                    } else {
                        self.compile_expr(object)?;
                        self.emit_autoderef_pointer_cell();
                        self.compile_array_index_operand_for_owner(object, index)?;
                        if self.profile.negative_index_wraps {
                            self.emit_negative_index_wrap();
                        }
                        {
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                    }
                }
            }

            // ── New ─────────────────────────────────────────────────────
            ExprKind::New { class, args } => {
                let reordered_args;
                let args = if args.iter().any(|arg| arg.name.is_some()) {
                    let ctor_key = match &class.kind {
                        ExprKind::Ident(name) => Some(self.canon(name)),
                        ExprKind::Member { field, .. } => Some(self.canon(field)),
                        _ => None };
                    if let Some(signatures) = ctor_key
                        .as_ref()
                        .and_then(|key| self.constructor_signatures.get(key))
                    {
                        reordered_args = self.reorder_named_args_with_signatures(args, signatures);
                        reordered_args.as_slice()
                    } else {
                        args.as_slice()
                    }
                } else {
                    args.as_slice()
                };
                // ECMA-262 §10.5.2: `new Proxy(target, handler)` creates an
                // exotic object whose property accesses are intercepted by
                // handler traps. We lower to an inline emitter that
                // produces a wrapper Ordinary object stamped with
                // `__vybe_proxy_target` + `__vybe_proxy_handler`. The
                // module-level `uses_proxy` flag (set by the AST scan)
                // tells subsequent Member/Index emits to route through
                // the dispatcher.
                if self.profile.ecma_new_dispatch {
                    if let ExprKind::Ident(name) = &class.kind {
                        if name == "Set" && args.len() <= 1 {
                            if let Some(arg) = args.first() {
                                self.compile_expr(&arg.value)?;
                                common::collections::emit_spread_iterable(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                            } else {
                                common::collections::emit_array_new(
                                    &mut self.chunks,
                                    self.current,
                                    0,
                                    self.line,
                                );
                            }
                            let idx = self.import("ecma:set", "fromIterable");
                            self.emit_host_call(idx, 1);
                            return Ok(());
                        }

                        if name == "Function" {
                            for arg in args {
                                self.compile_expr(&arg.value)?;
                            }
                            let idx = self.import("vybe:js", "function_constructor");
                            self.emit_host_call(idx, args.len() as u8);
                            return Ok(());
                        }

                        if name == "Promise" && args.len() == 1 {
                            self.compile_expr(&args[0].value)?;
                            let idx = self.import("ecma:promise", "new");
                            self.emit_host_call(idx, 1);
                            return Ok(());
                        }

                        if name == "Proxy" && args.len() == 2 {
                            self.uses_proxy = true;
                            self.compile_expr(&args[0].value)?;
                            self.compile_expr(&args[1].value)?;
                            let line = self.line;
                            vybe_runtime::registry::hooks(&self.profile.name)
                                .proxy_create
                                .unwrap()(
                                &mut self.chunks, self.current, line
                            );
                            return Ok(());
                        }
                    }
                }
                let class_parts = self.flatten_member_chain(class);
                let dotted_type_name = match &class.kind {
                    ExprKind::Ident(name) => {
                        let resolved = self.resolve_source_type_alias(name);
                        let canon_resolved = self.canon(&resolved);
                        if self.profile.uses_common_resolver
                            && self.canon(&resolved) == self.canon(name)
                        {
                            if let Some(qualified) = self
                                .current_namespace
                                .as_deref()
                                .map(|ns| self.canon(&format!("{ns}.{canon_resolved}")))
                                .filter(|qualified| self.defined_classes.contains(qualified))
                            {
                                Some(qualified)
                            } else {
                                Some(resolved)
                            }
                        } else {
                            Some(resolved)
                        }
                    }
                    ExprKind::Member { .. }
                        if self.profile.uses_namespace_resolver() && !class_parts.is_empty() =>
                    {
                        Some(self.resolve_source_type_alias(&class_parts.join(".")))
                    }
                    _ => None };
                // The source spelling to hand a runtime type resolver when the
                // constructor global turns out to be undefined. Only meaningful
                // where the language HAS such a resolver — see
                // `emit_constructor_global_ref`, which owns the dispatch.
                let autoload_source_name = match &class.kind {
                    ExprKind::Ident(name) if self.profile.supports_autoload => {
                        Some(Self::strip_global_namespace_prefix(name).to_string())
                    }
                    _ => None };

                if let Some(type_name) = dotted_type_name.as_ref() {
                    // User-defined classes take priority over all built-in type mappings.
                    // This ensures `class Point { ... }` followed by `new Point()` calls
                    // the user constructor, not vybe:drawing::pointNew.
                    let canon_type = self.canon(type_name);
                    if self.abstract_classes.contains(&canon_type) {
                        let line = self.line;
                        let chunk = self.chunk();
                        chunk.emit_struct_new(0, 0, line);
                        chunk.emit_dup(line);
                        chunk.emit_string_const(
                            &format!("Cannot instantiate abstract class {}", type_name),
                            line,
                        );
                        crate::primitives::errors::emit_exception_new_finalize(
                            chunk, "Error", line,
                        );
                        crate::primitives::errors::emit_throw(chunk, line);
                        return Ok(());
                    }
                    if self.defined_classes.contains(&canon_type) {
                        // §14.1.13 (JS): rest parameter in a constructor —
                        // pack surplus positional args into an Array so the
                        // ctor's rest slot receives a proper Array (no
                        // surplus ⇒ empty array). Static packing; spread
                        // args keep the plain path.
                        let js_ctor_rest_fixed: Option<usize> = if self.profile.ecma_new_dispatch
                            && !args
                                .iter()
                                .any(|a| matches!(a.value.kind, ExprKind::Spread(_)))
                        {
                            self.constructor_signatures
                                .get(&canon_type)
                                .and_then(|sigs| sigs.iter().find(|s| s.has_rest))
                                .map(|s| s.param_names.len().saturating_sub(1))
                                .filter(|fixed| args.len() >= *fixed)
                        } else {
                            None
                        };
                        let effective_len = match js_ctor_rest_fixed {
                            Some(fixed) => fixed + 1,
                            None => args.len() };
                        let overload_global =
                            crate::primitives::classes::ctor_global_for(&canon_type, effective_len);
                        let ctor_global = if self.defined_globals.contains(&overload_global) {
                            overload_global
                        } else {
                            canon_type.clone()
                        };
                        // Bypass compile_expr to avoid the implicit-self-field
                        // shadowing path: in case-insensitive languages a field
                        // named `inner` and a class named `Inner` both
                        // canonicalize to "inner", and the implicit-self-field
                        // check would mis-route to `me.inner` instead of the
                        // class global. Type names always come from globals.
                        let autoload_name = autoload_source_name.as_deref().unwrap_or(type_name);
                        self.emit_constructor_global_ref(&ctor_global, autoload_name);
                        if let Some(fixed) = js_ctor_rest_fixed {
                            for a in &args[..fixed] {
                                self.compile_expr(&a.value)?;
                            }
                            for a in &args[fixed..] {
                                self.compile_expr(&a.value)?;
                            }
                            let l = self.line;
                            common::collections::emit_array_new(
                                &mut self.chunks,
                                self.current,
                                (args.len() - fixed) as u16,
                                l,
                            );
                        } else {
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                        }
                        // §13.3.5: new.target is the invoked constructor for
                        // the whole construction chain (parent ctor bodies
                        // under super() included). Unique save slot — nested
                        // `new` in the same function must not share it.
                        let saved_nt = self.save_js_new_target(&format!(
                            "__js_prev_nt_static_{}",
                            self.chunks[self.current].local_count
                        ));
                        if saved_nt.is_some() {
                            self.emit_global_read(&canon_type);
                            self.set_js_new_target_from_stack();
                        }
                        let saved_this = self.save_js_this(&format!(
                            "__js_prev_this_new_{}",
                            self.chunks[self.current].local_count
                        ));
                        self.emit_u8(Op::CALL_REF, effective_len as u8);
                        self.restore_js_this(saved_this);
                        self.restore_js_new_target(saved_nt);
                        return Ok(());
                    }
                    if self.profile.ecma_new_dispatch
                        && self.defined_functions.contains(&canon_type)
                    {
                        self.emit_global_read(&canon_type);
                        let ctor_slot = self.define_local("__js_ctor");
                        self.emit_u16(Op::LOCAL_SET, ctor_slot);
                        let line = self.line;
                        let saved_js_new_target =
                            self.save_js_new_target("__js_prev_new_target_new");
                        self.emit_u16(Op::LOCAL_GET, ctor_slot);
                        self.set_js_new_target_from_stack();
                        let _ = line;
                        let (args_slot, _) = self.compile_call_args_array(args, "js_new")?;
                        self.emit_u16(Op::LOCAL_GET, ctor_slot);
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        common::reflection::emit_reflect_op(
                            &mut self.chunks,
                            self.current,
                            common::reflection::ReflectOp::Construct,
                            2,
                            line,
                        );
                        self.restore_js_new_target(saved_js_new_target);
                        return Ok(());
                    }
                    if self.profile.supports_autoload {
                        if let Some(autoload_name) = autoload_source_name.as_deref() {
                            if let Some(flattened_name) = autoload_name.rsplit('\\').next() {
                                let flattened_canon = self.canon(flattened_name);
                                if self.defined_classes.contains(&flattened_canon) {
                                    let overload_global =
                                        crate::primitives::classes::ctor_global_for(
                                            &flattened_canon,
                                            args.len(),
                                        );
                                    let ctor_global =
                                        if self.defined_globals.contains(&overload_global) {
                                            overload_global
                                        } else {
                                            flattened_canon
                                        };
                                    self.emit_constructor_global_ref(&ctor_global, autoload_name);
                                    for a in args {
                                        self.compile_expr(&a.value)?;
                                    }
                                    self.emit_u8(Op::CALL_REF, args.len() as u8);
                                    return Ok(());
                                }
                            }
                        }
                    }
                    // Nested class: `new Outer.Inner()` — the inner type
                    // is registered as a sibling global per ECMA-334 §15.3.
                    // Try the last segment as a type name when the full
                    // dotted form misses.
                    if class_parts.len() > 1 {
                        let last = class_parts.last().unwrap();
                        let canon_last = self.canon(last);
                        if self.defined_classes.contains(&canon_last) {
                            let autoload_name =
                                autoload_source_name.as_deref().unwrap_or(type_name);
                            self.emit_constructor_global_ref(&canon_last, autoload_name);
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            return Ok(());
                        }
                    }

                    let bare = type_name.to_lowercase();
                    let bare = bare.split('(').next().unwrap_or(&bare).trim();
                    let bare_str = bare.rsplit('.').next().unwrap_or(bare);

                    // WASM threading/async — use compiler_common, NOT host calls
                    match bare_str {
                        "thread" => {
                            // New Thread(callback) → cont_new only (Start resumes)
                            if let Some(a) = args.first() {
                                self.compile_expr(&a.value)?;
                            }
                            let line = self.line;
                            common::threading::emit_thread_new(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            return Ok(());
                        }
                        "task" => {
                            // New Task(callback) → cont_new only
                            if let Some(a) = args.first() {
                                self.compile_expr(&a.value)?;
                            }
                            let line = self.line;
                            common::threading::emit_thread_new(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            return Ok(());
                        }
                        "mutex" | "semaphore" => {
                            // New Mutex() → allocate atomic address for lock
                            self.emit_const(Value::I32(0)); // initial lock value
                            return Ok(());
                        }
                        _ => {}
                    }

                    // Built-in exception types with NO backing class definition —
                    // route through compiler_common so a language that treats its
                    // exceptions as intrinsics (JS `new Error()`, Python
                    // `RuntimeError`, …) still produces the canonical shape.
                    //
                    // The predicate is deliberately "built-in name AND not a
                    // defined class": once a language models its exceptions as
                    // real classes (PHP defines the whole Throwable/Error/
                    // Exception hierarchy), the name is user-visible and MUST go
                    // through the ordinary class emitter instead — that's what
                    // keeps `get_class` and the `__types` inheritance chain
                    // language-faithful. Not a special case: an intrinsic name
                    // shadowed by a real class is no longer an intrinsic.
                    let is_intrinsic_exception = common::errors::is_exception_type(bare_str)
                        && !self.defined_classes.contains(type_name)
                        && !self.defined_classes.contains(&self.canon(type_name))
                        && !self.defined_classes.contains(bare_str);
                    if is_intrinsic_exception {
                        let ctor_args: Vec<&Expression> = args.iter().map(|a| &a.value).collect();
                        self.emit_js_exception_ctor_value(type_name, &ctor_args)?;
                        return Ok(());
                    }

                    // Common-resolver construction (namespaceplan.md): a tree
                    // `Type` node carrying a `CtorSpec` — resolved via a mounted
                    // ambient root (`flutter.*`, …) — constructs generically,
                    // cross-language, through the ONE resolver. A user class of
                    // the same name shadows it.
                    if !self.defined_classes.contains(type_name)
                        && !self.defined_classes.contains(bare_str)
                        && !self.defined_classes.contains(&self.canon(type_name))
                    {
                        if let Some(super::resolver::Resolution::Tree(
                            crate::primitives::namespaces::ResolutionTarget::Ctor {
                                spec: Some(spec),
                                ..
                            },
                        )) = self.resolve_profile_namespace_chain(&[type_name.to_string()])
                        {
                            // Generic field-capture construction only when the
                            // spec actually DESCRIBES a construction — captured
                            // params/fields, or a `vybe:gui` control factory
                            // (flutter widgets, plib GCL controls).
                            //
                            // A spec may instead be IDENTITY-ONLY: just an
                            // `ancestry`, declared so `isInstance`/`is` can
                            // answer, with the real constructor held in the
                            // type's `ctor_call` (Java's stdlib types are this
                            // shape — `new ArrayList()` is
                            // `common:java.mutable_list_of`). Building those
                            // generically yields an EMPTY object and loses the
                            // constructor entirely, so they fall through to the
                            // backing-call path below, which also stamps the
                            // ancestry.
                            let describes_construction = !spec.params.is_empty()
                                || !spec.fields.is_empty()
                                || spec.control_fn.is_some();
                            if describes_construction {
                                return self.emit_tree_ctor_construction(&spec, args);
                            }
                        }
                    }

                    // GUI control: Button, TextBox, Label, Timer, etc.
                    // Checked BEFORE dotnet known_types so GUI controls always
                    // route through the canonical gui emitter regardless of
                    // whether they overlap with .NET BCL types (Timer is both
                    // a GUI control and a System.Threading.Timer — the GUI
                    // form takes priority because we're in `New X()` syntax).
                    // A user global spelling a control name shadows the control
                    // factory — the same rule `is_framework_control_parent`
                    // already applies for a user CLASS. That is a fact about
                    // the program, not about which language family compiled it.
                    let dotnet_ctor_registered = self.defined_globals.contains(bare_str)
                        || self.defined_globals.contains(&bare_str.to_lowercase());
                    let canonical = common::gui::canonical_control_name(bare_str);
                    if !canonical.is_empty() && !dotnet_ctor_registered {
                        for a in args {
                            self.compile_expr(&a.value)?;
                        }
                        let line = self.line;
                        self.emit_control_element(bare_str, args.len() as u8, line);
                        return Ok(());
                    }
                    // Registered-type constructors — after GUI so .NET-only
                    // types like Dictionary still work.
                    //
                    // A language-name fork used to sit here, sending one
                    // language's `new X()` straight to `lookup_known_type` and
                    // returning. That skipped the identity stamp below, so
                    // every stdlib type that language registered was
                    // constructible but anonymous — no `__type`, no `__types` —
                    // and the language grew its own `isInstance` fallback to
                    // compensate. Types the tree does not register still reach
                    // `lookup_known_type` at the fallback further down, so the
                    // only behaviour change is that registered types now carry
                    // their declared ancestry.
                    let dotnet_constructor = vybe_runtime::namespaces::lookup_type_ctor_target(
                        &self.profile.namespaces.type_scopes,
                        bare_str,
                    );
                    if let Some(target) = dotnet_constructor.clone() {
                        for a in args {
                            self.compile_expr(&a.value)?;
                        }
                        // Proper-case class name (preserve from source) for
                        // the __type stamp. Some host fns (e.g.
                        // `vybe:types/collectionPeek`) compare __type with
                        // exact case, so a lowercase stamp would clobber.
                        let proper_name: String = type_name
                            .split('(')
                            .next()
                            .unwrap_or(type_name)
                            .trim()
                            .rsplit('.')
                            .next()
                            .unwrap_or(type_name)
                            .to_string();
                        match target {
                            vybe_runtime::component_model::ConstructorTarget::Host(target) => {
                                let idx = self.import(&target.module, &target.name);
                                self.emit_host_call(idx, args.len() as u8);
                            }
                            vybe_runtime::component_model::ConstructorTarget::Common(name) => {
                                let line = self.line;
                                self.emit_common(&name, args.len() as u8, line);
                            }
                        }
                        // Stamp `__type` with the .NET class name so the
                        // runtime TypeRegistry dispatches `d.Add(...)` /
                        // `d.ContainsKey(...)` against the .NET adapter
                        // TypeDef (e.g. `Dictionary`) — not against the
                        // underlying ECMA type (e.g. `Map`). The .NET
                        // adapter's methods are aliases pointing at the
                        // same `ecma:*` host fns, so the underlying
                        // implementation is standardized while the
                        // surface stays .NET-shaped.
                        //
                        // Stack: [obj] → [obj, obj, name] → [obj, obj'] → [obj]
                        inst!(self, core_wasm::dup);
                        self.emit_const(Value::String(Arc::from(proper_name.as_str())));
                        let type_key = self.str_const("__type");
                        self.emit_struct_field_op(Op::STRUCT_SET, 0, type_key);

                        // …and stamp the ANCESTRY chain when the registered
                        // type declares one, so `isInstance` / `instanceof` /
                        // `is` answer from the shared `__types` array
                        // (`reflection::emit_instanceof`) exactly as they do
                        // for a user class.
                        //
                        // Without this a platform type was constructible but
                        // not IDENTIFIABLE: `new ArrayList()` produced a value
                        // with no `__types`, so the shared check could not
                        // answer and each language grew its own fallback — Java
                        // carried a hardcoded ~30-name list plus a
                        // `__java_class_is_instance` helper for exactly this.
                        // The ancestry is data the platform already declares.
                        if let Some(spec) = vybe_runtime::namespaces::lookup_type_ctor_spec(
                            &self.profile.namespaces.type_scopes,
                            bare_str,
                        ) {
                            if !spec.ancestry.is_empty() {
                                let line = self.line;
                                inst!(self, core_wasm::dup);
                                for ancestor in &spec.ancestry {
                                    self.emit_const(Value::String(Arc::from(ancestor.as_str())));
                                }
                                crate::primitives::collections::emit_array_new(
                                    &mut self.chunks,
                                    self.current,
                                    spec.ancestry.len() as u16,
                                    line,
                                );
                                let types_key =
                                    self.str_const(crate::primitives::reflection::FIELD_TYPES);
                                self.emit_struct_field_op(Op::STRUCT_SET, 0, types_key);
                            }
                        }

                        // .NET List / ArrayList instance calls like `list.Sort()`
                        // should stay on the shared compare-aware frontend path
                        // instead of falling through to the runtime collection
                        // vtable's raw host sort mapping.
                        let bare_proper_name = proper_name
                            .split('<')
                            .next()
                            .map(str::trim)
                            .unwrap_or(proper_name.as_str());
                        if matches!(bare_proper_name, "List" | "ArrayList") {
                            let sort_key = self.str_const("sort");
                            inst!(self, core_wasm::dup);
                            self.emit_global_read("__vybe_sort_in_place");
                            self.emit_struct_field_op(Op::STRUCT_SET, 0, sort_key);

                            let sort_pascal_key = self.str_const("Sort");
                            inst!(self, core_wasm::dup);
                            self.emit_global_read("__vybe_sort_in_place");
                            self.emit_struct_field_op(Op::STRUCT_SET, 0, sort_pascal_key);
                        }
                        return Ok(());
                    }

                    // Profile known types are now a fallback for entries not
                    // yet absorbed into the shared .NET surface.
                    if dotnet_constructor.is_none() {
                        if let Some((module, func)) = self
                            .profile
                            .lookup_known_type(type_name)
                            .map(|(m, f)| (m.to_string(), f.to_string()))
                        {
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            if module == "common" {
                                let line = self.line;
                                self.emit_common(&func, args.len() as u8, line);
                            } else {
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, args.len() as u8);
                                // §20.5: JS error instances — pure-WASM
                                // prototype link + own-name removal.
                                if self.profile.ecma_new_dispatch && module == "ecma:error" {
                                    let line = self.line;
                                    crate::primitives::errors::emit_finish_js_error_instance(
                                        self.chunk(),
                                        &func,
                                        line,
                                    );
                                }
                            }
                            return Ok(());
                        }
                    }

                    if dotnet_ctor_registered {
                        let ctor_name = if self.defined_globals.contains(bare_str) {
                            bare_str.to_string()
                        } else {
                            bare_str.to_lowercase()
                        };
                        let autoload_name = autoload_source_name.as_deref().unwrap_or(type_name);
                        self.emit_constructor_global_ref(&ctor_name, autoload_name);
                        for a in args {
                            self.compile_expr(&a.value)?;
                        }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);

                        if bare_str.eq_ignore_ascii_case("list")
                            || bare_str.eq_ignore_ascii_case("arraylist")
                        {
                            let sort_key = self.str_const("sort");
                            inst!(self, core_wasm::dup);
                            self.emit_global_read("__vybe_sort_in_place");
                            self.emit_struct_field_op(Op::STRUCT_SET, 0, sort_key);

                            let sort_pascal_key = self.str_const("Sort");
                            inst!(self, core_wasm::dup);
                            self.emit_global_read("__vybe_sort_in_place");
                            self.emit_struct_field_op(Op::STRUCT_SET, 0, sort_pascal_key);
                        }
                        return Ok(());
                    }

                    if self.profile.supports_autoload {
                        // An Ident in the VARIABLE namespace is a variable
                        // (`new $c`), not a class name — its runtime string
                        // value is resolved to a constructor by the dynamic
                        // fall-through below.
                        if let ExprKind::Ident(name) = &class.kind {
                            if !self.is_variable_name(name) {
                                let autoload_name =
                                    Self::strip_global_namespace_prefix(name).to_string();
                                let ctor_base = autoload_name
                                    .rsplit('\\')
                                    .next()
                                    .unwrap_or(autoload_name.as_str());
                                let fallback_ctor = self.canon(ctor_base);
                                let primary_ctor = crate::primitives::classes::ctor_global_for(
                                    &fallback_ctor,
                                    args.len(),
                                );
                                self.emit_dynamic_constructor_global_ref(
                                    &primary_ctor,
                                    Some(&fallback_ctor),
                                    &autoload_name,
                                );
                                for a in args {
                                    self.compile_expr(&a.value)?;
                                }
                                self.emit_u8(Op::CALL_REF, args.len() as u8);
                                return Ok(());
                            }
                        }
                    }
                }
                if self.profile.ecma_new_dispatch {
                    self.compile_expr(class)?;
                    let ctor_slot = self.define_local("__js_ctor");
                    self.emit_u16(Op::LOCAL_SET, ctor_slot);
                    let saved_js_new_target = self.save_js_new_target("__js_prev_new_target_new");
                    self.emit_u16(Op::LOCAL_GET, ctor_slot);
                    self.set_js_new_target_from_stack();
                    let (args_slot, _) = self.compile_call_args_array(args, "js_new")?;
                    self.emit_u16(Op::LOCAL_GET, ctor_slot);
                    self.emit_u16(Op::LOCAL_GET, args_slot);
                    common::reflection::emit_reflect_op(
                        &mut self.chunks,
                        self.current,
                        common::reflection::ReflectOp::Construct,
                        2,
                        self.line,
                    );
                    self.restore_js_new_target(saved_js_new_target);
                    return Ok(());
                }
                self.compile_expr(class)?;
                for a in args {
                    self.compile_expr(&a.value)?;
                }
                self.emit_u8(Op::CALL_REF, args.len() as u8);
            }

            // ── Assignment as expression ────────────────────────────────
            ExprKind::Assign { target, value } => {
                if matches!(value.kind, ExprKind::Lit(Literal::Null)) {
                    if let ExprKind::Ident(name) = &target.kind {
                        self.emit_buffered_generator_close_ident_if_needed(name);
                    }
                }
                self.compile_expr(value)?;
                inst!(self, core_wasm::dup);
                self.compile_assign_target_valued(target, Some(value))?;
                // Reference assignment: mark the target as a pointer-cell
                // binding AFTER the first store, so subsequent writes go
                // THROUGH the reference instead of overwriting the name.
                //
                // Both spellings of "make a reference" count. `RefOf` and
                // `Unary{AddrOf}` are the same concept (§2 of `referenceplan.md`)
                // and matching only one made the binding's aliasing depend on
                // which node the walker happened to build — a name bound to a
                // reference silently stayed a plain value.
                if matches!(
                    &value.kind,
                    ExprKind::RefOf(_)
                        | ExprKind::Unary {
                            op: UnaryOp::AddrOf,
                            ..
                        }
                ) {
                    if let ExprKind::Ident(name) = &target.kind {
                        self.mark_pointer_cell_binding(name);
                    }
                }
            }

            // ── Lambda ──────────────────────────────────────────────────
            ExprKind::Lambda {
                params,
                body,
                captures,
                is_async } => {
                // ExprKind::Lambda in JS IS the arrow form (function
                // expressions arrive as FunctionExpr, shorthand methods
                // via the object-literal path below).
                self.compile_lambda_with_flags(params, body, captures, *is_async, false, true)?;
            }

            // ── Array literal ───────────────────────────────────────────
            ExprKind::Array(elements) => {
                // Array literals funnel through `common::collections` so
                // every language emits the same import shape.
                //
                // Dispatch on whether ANY element has an explicit key:
                //   - no keys  → `ecma:array` path (integer-indexed,
                //                fast, array-y semantics)
                //   - any key  → `ecma:map` path (IndexMap<Value,Value>,
                //                PHP-shaped associative: mixed int + string
                //                keys preserved in insertion order). Once
                //                a Map, accessors `$a[$k]` use
                //                `ecma:array.get/.set` which now
                //                dispatch polymorphically on Map.
                let line = self.line;
                let allows_array_elisions = self.profile.ecma_array_elisions;
                let is_array_elision = |elem: &ArrayElement| {
                    allows_array_elisions
                        && !elem.spread
                        && matches!(&elem.key, Some(key) if matches!(key.kind, ExprKind::Lit(Literal::Int(-1))))
                        && matches!(elem.value.kind, ExprKind::Lit(Literal::Undefined))
                };
                let has_keys = elements
                    .iter()
                    .any(|e| e.key.is_some() && !is_array_elision(e));
                let has_elisions = allows_array_elisions && elements.iter().any(is_array_elision);

                if matches!(self.profile.name.as_str(), "c" | "vb")
                    && !has_keys
                    && !has_elisions
                    && elements.iter().all(|elem| !elem.spread)
                {
                    for elem in elements {
                        self.compile_expr(&elem.value)?;
                    }
                    self.emit_array_new_fixed(0, elements.len() as u16);
                    return Ok(());
                }

                if has_keys {
                    common::collections::emit_map_new(&mut self.chunks, self.current, line);
                    let mut next_auto_idx: i64 = 0;
                    for elem in elements {
                        if elem.spread {
                            // Spread into a keyed literal isn't
                            // meaningful for Maps in PHP/Python/Ruby
                            // semantics in Phase 1 — skip.
                            continue;
                        }
                        // Stack: [map]
                        inst!(self, core_wasm::dup); // [map, map]
                        match &elem.key {
                            Some(k) => {
                                // Explicit key expression. If it's a
                                // numeric literal, bump the auto index
                                // past it so subsequent unkeyed elements
                                // don't collide (PHP semantics).
                                self.compile_expr(k)?;
                                match &k.kind {
                                    ExprKind::Lit(crate::ast::Literal::Int(n)) => {
                                        next_auto_idx = *n + 1;
                                    }
                                    ExprKind::Lit(crate::ast::Literal::Float(n))
                                        if n.fract() == 0.0 =>
                                    {
                                        next_auto_idx = (*n as i64) + 1;
                                    }
                                    _ => {}
                                }
                            }
                            None => {
                                // Unkeyed element inside a keyed literal
                                // (PHP allows `['x' => 1, 'y']`) — auto
                                // assign the next integer key.
                                self.emit_const(Value::I32(next_auto_idx as i32));
                                next_auto_idx += 1;
                            }
                        }
                        // [map, map, key]
                        self.compile_expr(&elem.value)?;
                        // [map, map, key, value]
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        // set returns null per spec — drop it, map stays on TOS
                        self.emit(Op::DROP);
                    }
                } else {
                    if has_elisions {
                        let array_new_idx = self.import("ecma:array", "new");
                        self.emit_const(Value::I32(elements.len() as i32));
                        self.emit_host_call(array_new_idx, 1);
                        for (index, elem) in elements.iter().enumerate() {
                            if is_array_elision(elem) {
                                continue;
                            }
                            inst!(self, core_wasm::dup);
                            self.emit_const(Value::I32(index as i32));
                            self.compile_expr(&elem.value)?;
                            common::collections::emit_set(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                        }
                    } else {
                        // All-unkeyed: use the array path (fast, small).
                        common::collections::emit_array_new(
                            &mut self.chunks,
                            self.current,
                            0,
                            line,
                        );
                        for elem in elements {
                            if elem.spread {
                                // Spread: `concat(current, other)` returns a NEW
                                // array which replaces the one on TOS. JS
                                // generators (Continuation values) can't be
                                // spread by the host concat fn — the iterator
                                // protocol needs spec stack-switching `resume`
                                // (emit_next). When we know at compile time that
                                // the spread value is a direct generator call,
                                // use the common emitter (emit_drain_into_array)
                                // which emits an inline GEN_NEXT loop — the same
                                // opcode compile_generator_for_in uses directly.
                                // No runtime helper, no runtime dispatch, no isGenerator check.
                                // For unknown-shape values, fall back to the runtime
                                // isGenerator dispatch via stdlib.
                                // A direct call to a generator function can be
                                // drained at compile time via the common GEN_NEXT
                                // loop (WASM stack switching) — works for ANY
                                // language with generators (JS, PHP), no runtime
                                // isGenerator check, no runtime helper.
                                let is_known_gen = (self.profile.has_generators
                                    || self.profile.buffered_iterator_methods)
                                    && self.is_direct_generator_call(&elem.value);
                                self.compile_expr(&elem.value)?;
                                if is_known_gen {
                                    // continuation is on TOS; drain into array
                                    crate::primitives::generators::emit_drain_into_array(
                                        &mut self.chunks,
                                        self.current,
                                        line,
                                    );
                                    common::collections::emit_concat(
                                        &mut self.chunks,
                                        self.current,
                                        line,
                                    );
                                    continue;
                                }
                                common::collections::emit_spread_iterable(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                common::collections::emit_concat(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                            } else {
                                // DUP keeps the array on TOS; push returns the
                                // new length, which we drop.
                                inst!(self, core_wasm::dup);
                                self.compile_expr(&elem.value)?;
                                common::collections::emit_push(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                self.emit(Op::DROP);
                            }
                        }
                    }
                }
            }

            // ── Cross-language zip primitive ───────────────────────────
            ExprKind::Zip {
                iterables,
                mode,
                strict } => {
                let line = self.line;
                if *strict {
                    return Err("strict zip lowering is not wired yet".into());
                }
                for iterable in iterables {
                    self.compile_expr(iterable)?;
                }
                let mode = match mode {
                    crate::ast::ZipMode::First => common::collections::ZipLen::First,
                    crate::ast::ZipMode::Shortest => common::collections::ZipLen::Shortest,
                    crate::ast::ZipMode::Longest => common::collections::ZipLen::Longest };
                common::collections::emit_zip(
                    &mut self.chunks,
                    self.current,
                    iterables.len() as u8,
                    mode,
                    line,
                );
            }

            // ── Tuple (Python) ──────────────────────────────────────────
            ExprKind::Tuple(elements) => {
                let line = self.line;
                let n = elements.len();
                for elem in elements {
                    self.compile_expr(elem)?;
                }
                // Allocate N consecutive slots; common::collections::emit_pack_n
                // stashes stack values and re-pushes into a fresh array —
                // same ecma:array.* surface as literals.
                let base = if n == 0 {
                    0
                } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.define_local("__pack");
                        if i == 0 {
                            first = s;
                        }
                    }
                    first
                };
                common::collections::emit_pack_n(
                    &mut self.chunks,
                    self.current,
                    n as u16,
                    base,
                    line,
                );
                // A tuple keeps the array as its underlying value but is tagged
                // so repr/type()/slicing distinguish it from a list. Shared,
                // cross-language: opt in via the `tuple_literals_tagged` profile.
                if self.profile.tuple_literals_tagged {
                    common::tuples::emit_tag(&mut self.chunks, self.current, line);
                }
            }

            // ── Named tuple (C# `(x: 1, y: 2)`, Python `namedtuple`) ─────
            // Same tagged-array backing as a plain tuple, then the shared
            // emitter stamps by-name field keys + `__fields`/`__typename`.
            ExprKind::NamedTuple { fields, type_name } => {
                let line = self.line;
                let n = fields.len();
                for (_, value) in fields {
                    self.compile_expr(value)?;
                }
                let base = if n == 0 {
                    0
                } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.define_local("__pack");
                        if i == 0 {
                            first = s;
                        }
                    }
                    first
                };
                common::collections::emit_pack_n(
                    &mut self.chunks,
                    self.current,
                    n as u16,
                    base,
                    line,
                );
                let names: Vec<Option<String>> =
                    fields.iter().map(|(name, _)| name.clone()).collect();
                common::tuples::emit_named_tuple(
                    &mut self.chunks,
                    self.current,
                    &names,
                    type_name.as_deref(),
                    line,
                );
            }

            // ── Set (Python) ────────────────────────────────────────────
            ExprKind::Set(elements) => {
                let line = self.line;
                let n = elements.len();
                for elem in elements {
                    self.compile_expr(elem)?;
                }
                let base = if n == 0 {
                    0
                } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.define_local("__pack");
                        if i == 0 {
                            first = s;
                        }
                    }
                    first
                };
                common::collections::emit_pack_n(
                    &mut self.chunks,
                    self.current,
                    n as u16,
                    base,
                    line,
                );
                // Convert the packed array to a Set per ECMA-262 §24.2.1.1
                // — `new Set(iterable)`. Same import V8 satisfies natively
                // for `new Set([1,2,3])`.
                let idx = self.import("ecma:set", "fromIterable");
                self.emit_host_call(idx, 1);
            }

            // ── Object literal ──────────────────────────────────────────
            // Uses dict::emit_new to create the object WITH __keys tracking.
            // Each key is set via struct_set AND appended to __keys so that
            // Object.keys/values/entries (which read __keys) return the
            // right answer.
            // An ordered, `Value`-keyed literal. Build `[[k, v], …]` then
            // `Map.fromEntries` — keeps key types and insertion order.
            //
            // No profile consulted: the NODE says it is a Map. The branch that
            // used to live inside `ExprKind::Object` asked
            // `profile.dict_literals_as_map`, so the same node compiled to two
            // different runtime shapes depending on which language emitted it —
            // and a primitive holding one could not tell which it had.
            ExprKind::Map(entries) => {
                for (key, value) in entries {
                    self.compile_expr(key)?;
                    self.compile_expr(value)?;
                    self.emit_array_new_fixed(0, 2);
                }
                self.emit_array_new_fixed(0, entries.len() as u16);
                let idx = self.import("ecma:map", "fromEntries");
                self.emit_host_call(idx, 1);
            }

            ExprKind::Object(props) => {
                let line = self.line;
                common::dict::emit_new(&mut self.chunks, self.current, line);
                for prop in props {
                    match prop {
                        ObjectProperty::KeyValue { key, value } => {
                            if let ExprKind::Lit(Literal::Str(k)) = &key.kind {
                                // Static string key — fast path via struct_set.
                                let key_name = if self.case_sensitive {
                                    k.clone()
                                } else {
                                    self.canon(k)
                                };
                                // Track insertion order BEFORE setting the
                                // property. `trackKey` dedups by skipping keys
                                // already present — so `{ ...base, x: 99 }`
                                // (where `x` arrived via the spread) does NOT
                                // append a duplicate `x` to `__keys`. Must run
                                // pre-set: post-set the property always exists
                                // and tracking would be skipped entirely.
                                if self.profile.ecma_object_literals {
                                    inst!(self, core_wasm::dup);
                                    self.emit_const(Value::String(Arc::from(key_name.as_str())));
                                    let track_idx = self.import("ecma:object", "trackKey");
                                    self.emit_host_call(track_idx, 2);
                                    self.emit(Op::DROP);
                                }
                                inst!(self, core_wasm::dup);
                                self.compile_expr(value)?;
                                if self.profile.ecma_object_literals {
                                    let should_infer_name = match &value.kind {
                                        ExprKind::Lambda { .. } => true,
                                        ExprKind::FunctionExpr(stmt) => {
                                            // Walker-synthesized `__anon_fn_N`
                                            // names are anonymous for §10.2.9
                                            // SetFunctionName purposes.
                                            matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name.is_empty() || name.starts_with("__anon_fn_"))
                                        }
                                        ExprKind::ClassExpr { name, .. } => name.is_none(),
                                        _ => false };
                                    if should_infer_name {
                                        let inferred_name = if let Some(stripped) =
                                            key_name.strip_prefix("__get_")
                                        {
                                            format!("get {}", stripped)
                                        } else if let Some(stripped) =
                                            key_name.strip_prefix("__set_")
                                        {
                                            format!("set {}", stripped)
                                        } else {
                                            key_name.clone()
                                        };
                                        inst!(self, core_wasm::dup);
                                        self.emit_const(Value::String(Arc::from(
                                            inferred_name.as_str(),
                                        )));
                                        let name_key = self.str_const("name");
                                        self.emit_struct_field_op(Op::STRUCT_SET, 0, name_key);
                                    }
                                }
                                let idx = self.str_const(&key_name);
                                self.emit_struct_field_op(Op::STRUCT_SET, 0, idx);
                                // Non-JS: append to __keys directly (JS already
                                // tracked it via the deduping `trackKey` above).
                                if !self.profile.ecma_object_literals {
                                    inst!(self, core_wasm::dup);
                                    let keys_key = self.str_const("__keys");
                                    self.emit_struct_field_op(Op::STRUCT_GET, 0, keys_key);
                                    self.emit_const(Value::String(Arc::from(key_name.as_str())));
                                    let l = self.line;
                                    common::collections::emit_push(
                                        &mut self.chunks,
                                        self.current,
                                        l,
                                    );
                                    self.emit(Op::DROP);
                                }
                            } else {
                                // Dynamic key — emit_set is
                                // `ecma:array.set(obj, key, value) → null`
                                // so we must push key BEFORE value. The
                                // previous impl pushed value then key,
                                // causing `{ 1: "one" }` to be stored
                                // under key "one" with value 1. Fix
                                // matches the canonical emit_set contract.
                                inst!(self, core_wasm::dup); // [dict, dict]
                                self.compile_expr(key)?; // [dict, dict, key]
                                inst!(self, core_wasm::dup); // [dict, dict, key, key]
                                let key_tmp = self.define_local("__obj_dyn_key");
                                self.emit_u16(Op::LOCAL_SET, key_tmp);
                                // [dict, dict, key]
                                self.compile_expr(value)?; // [dict, dict, key, value]
                                let l = self.line;
                                common::collections::emit_set(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP); // drop returned null
                                // Track dynamic key in __keys (stringified)
                                inst!(self, core_wasm::dup);
                                let keys_key = self.str_const("__keys");
                                self.emit_struct_field_op(Op::STRUCT_GET, 0, keys_key);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                let l = self.line;
                                common::collections::emit_push(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP);
                            }
                        }
                        ObjectProperty::Shorthand(name) => {
                            inst!(self, core_wasm::dup);
                            self.emit_var_get(name);
                            let idx = self.str_const(name);
                            self.emit_struct_field_op(Op::STRUCT_SET, 0, idx);
                            // Track key in __keys
                            inst!(self, core_wasm::dup);
                            let keys_key = self.str_const("__keys");
                            self.emit_struct_field_op(Op::STRUCT_GET, 0, keys_key);
                            self.emit_const(Value::String(Arc::from(name.as_str())));
                            let l = self.line;
                            common::collections::emit_push(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Spread(expr) => {
                            // Object spread mutates the in-progress object and must
                            // leave that object on the stack for the remaining literal.
                            let target_tmp = self.define_local("__obj_spread_target");
                            self.emit_u16(Op::LOCAL_SET, target_tmp);
                            self.emit_u16(Op::LOCAL_GET, target_tmp);
                            self.compile_expr(expr)?;
                            let idx = self.import("ecma:object", "assign");
                            self.emit_host_call(idx, 2);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, target_tmp);
                        }
                        ObjectProperty::Method { key, value } => {
                            inst!(self, core_wasm::dup);
                            if let StmtKind::FunctionDecl {
                                params,
                                body,
                                is_generator,
                                is_async,
                                ..
                            } = &value.kind
                            {
                                if self.profile.ecma_object_literals {
                                    self.compile_lambda_with_flags(
                                        params,
                                        &LambdaBody::Block(body.clone()),
                                        &[],
                                        *is_async,
                                        *is_generator,
                                        false,
                                    )?;
                                } else {
                                    // Object methods receive `this` as implicit first arg
                                    let mut method_params = vec![Param {
                                        name: self.profile.self_keyword.clone(),
                                        type_hint: None,
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: false,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false }];
                                    method_params.extend(params.iter().cloned());
                                    self.compile_lambda_with_flags(
                                        &method_params,
                                        &LambdaBody::Block(body.clone()),
                                        &[],
                                        *is_async,
                                        *is_generator,
                                        false,
                                    )?;
                                }
                            } else {
                                self.emit_null();
                            }
                            // Set fn.name = key for Function.prototype.name support.
                            if self.profile.ecma_object_literals {
                                inst!(self, core_wasm::dup);
                                self.emit_const(Value::String(Arc::from(key.as_str())));
                                let name_key = self.str_const("name");
                                self.emit_struct_field_op(Op::STRUCT_SET, 0, name_key);
                            }
                            let idx = self.str_const(key);
                            self.emit_struct_field_op(Op::STRUCT_SET, 0, idx);
                        }
                        ObjectProperty::Accessor { kind, key, value } => {
                            inst!(self, core_wasm::dup);
                            if let StmtKind::FunctionDecl {
                                params,
                                body,
                                is_generator,
                                is_async,
                                ..
                            } = &value.kind
                            {
                                if self.profile.ecma_object_literals {
                                    self.compile_lambda_with_flags(
                                        params,
                                        &LambdaBody::Block(body.clone()),
                                        &[],
                                        *is_async,
                                        *is_generator,
                                        false,
                                    )?;
                                } else {
                                    // Accessors receive `this` as first arg
                                    let mut accessor_params = vec![Param {
                                        name: self.profile.self_keyword.clone(),
                                        type_hint: None,
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: false,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false }];
                                    accessor_params.extend(params.iter().cloned());
                                    self.compile_lambda_with_flags(
                                        &accessor_params,
                                        &LambdaBody::Block(body.clone()),
                                        &[],
                                        *is_async,
                                        *is_generator,
                                        false,
                                    )?;
                                }
                            } else {
                                self.emit_null();
                            }
                            let accessor_name = match kind {
                                AccessorKind::Get => format!("get {}", key),
                                AccessorKind::Set => format!("set {}", key) };
                            inst!(self, core_wasm::dup);
                            self.emit_const(Value::String(Arc::from(accessor_name.as_str())));
                            let name_key = self.str_const("name");
                            self.emit_struct_field_op(Op::STRUCT_SET, 0, name_key);
                            let accessor_slot = match kind {
                                AccessorKind::Get => format!("__get_{}", key),
                                AccessorKind::Set => format!("__set_{}", key) };
                            let idx = self.str_const(&accessor_slot);
                            self.emit_struct_field_op(Op::STRUCT_SET, 0, idx);
                        }
                        ObjectProperty::Computed { key, value } => {
                            // ecma:array.set expects [obj, key, val] → null
                            inst!(self, core_wasm::dup);
                            self.compile_expr(key)?;
                            inst!(self, core_wasm::dup); // save key for trackKey
                            let key_tmp = self.define_local("__obj_comp_key");
                            self.emit_u16(Op::LOCAL_SET, key_tmp);
                            self.compile_expr(value)?;
                            // §10.2.9 SetFunctionName: an anonymous fn under
                            // a computed key is named from the runtime key
                            // (symbols → "[<description>]").
                            if self.profile.ecma_object_literals
                                && matches!(
                                    &value.kind,
                                    ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_)
                                )
                            {
                                inst!(self, core_wasm::dup);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                let sfn = self.import("ecma:function", "setFunctionName");
                                self.emit_host_call(sfn, 2);
                                self.emit(Op::DROP);
                            }
                            let l = self.line;
                            common::collections::emit_set(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP); // drop returned null
                            // Track key — host fn checks if it's a
                            // Symbol and routes to `__sym_keys` so
                            // Object.keys excludes it (ECMA-262 §7.3.22).
                            inst!(self, core_wasm::dup);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            let track_idx = self.import("ecma:object", "trackKey");
                            self.emit_host_call(track_idx, 2);
                            self.emit(Op::DROP);
                        }
                    }
                }
            }

            // ── String interpolation ────────────────────────────────────
            ExprKind::Interpolation(parts) => {
                if parts.is_empty() {
                    self.emit_const(Value::String(Arc::from("")));
                    return Ok(());
                }
                // For JS, ECMA-262 §13.2.8 specifies template literal
                // substitutions go through GetValue → ToString. ToString
                // on Objects calls ToPrimitive(hint=string) which prefers
                // toString over valueOf — exactly what the
                // `__vybe_to_primitive` polyfill implements. Routing
                // through it ensures user `toString` overrides fire
                // (sets `__js_this` via the JS method-call protocol),
                // then `"" + primitive` produces the final string.
                let use_to_primitive = self.profile.ecma_to_primitive;
                // How ONE substitution becomes a string is the LANGUAGE's rule,
                // not a shared one: PHP renders `null` and `false` as `""` and
                // `true` as `"1"`, Ruby calls `to_s`, JS follows ToString. This
                // arm already conceded that with `ecma_to_primitive` — a bool
                // for one language — so the general form is the binding.
                //
                // Same `[builtin_slots.string] to_string` the `.`/`..` concat
                // operator reads (§3i). Interpolation was hardcoded to
                // `strings::emit_to_string`, which is the ECMA coercion, so
                // `"{$x}"` on a null printed `null` and on a missing array key
                // `undefined` where PHP prints nothing.
                let interp_to_string = self
                    .profile
                    .builtin_slots
                    .get(
                        vybe_ast::builtin_slots::BuiltinType::String,
                        vybe_ast::ProtocolSlot::ToString,
                    )
                    .map(str::to_string);
                self.emit_const(Value::String(Arc::from("")));
                let acc_slot = self.define_local("__interp_acc");
                self.emit_u16(Op::LOCAL_SET, acc_slot);
                let part_slot = self.define_local("__interp_part");

                for part in parts.iter() {
                    match part {
                        InterpolPart::Text(s) => {
                            self.emit_const(Value::String(Arc::from(s.as_str())))
                        }
                        InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => {
                            if use_to_primitive {
                                self.compile_expr(e)?;
                                self.emit_to_primitive("string");
                                // After ToPrimitive, the value is a primitive
                                // (string / number / symbol / …). ECMA-262
                                // §13.2.8.6 evaluates each substitution with
                                // ToString (§7.1.17), which — unlike the
                                // `String(sym)` constructor — throws a TypeError
                                // for a Symbol. Guard for that before the concat.
                                let v_slot = self.define_local("__interp_to_string");
                                self.emit_u16(Op::LOCAL_SET, v_slot);
                                self.emit_u16(Op::LOCAL_GET, v_slot);
                                let sym_test = self.import("wasm:js-symbol", "test");
                                self.emit_host_call(sym_test, 1);
                                let line = self.line;
                                self.chunk().emit_if(line);
                                self.emit_const(Value::String(Arc::from(
                                    "Cannot convert a Symbol value to a string",
                                )));
                                self.emit_js_exception_ctor_from_message_value("TypeError")?;
                                common::errors::emit_throw(self.chunk(), line);
                                self.chunk().emit_end(line);
                                // Primitive -> string after the template-literal
                                // Symbol guard above.
                                self.emit_u16(Op::LOCAL_GET, v_slot);
                                let line = self.line;
                                common::strings::emit_to_string(self.chunk(), line);
                            } else {
                                self.compile_expr(e)?;
                                let value_slot = self.define_local("__interp_value");
                                self.emit_u16(Op::LOCAL_SET, value_slot);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                let line = self.line;
                                match &interp_to_string {
                                    Some(target) => {
                                        self.emit_slot_target(target, 1, line, "string to_string")
                                    }
                                    None => common::strings::emit_to_string(self.chunk(), line) }
                            }
                        }
                    }

                    self.emit_u16(Op::LOCAL_SET, part_slot);
                    self.emit_u16(Op::LOCAL_GET, acc_slot);
                    self.emit_u16(Op::LOCAL_GET, part_slot);
                    let line = self.line;
                    common::strings::emit_str_concat(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, acc_slot);
                }

                self.emit_u16(Op::LOCAL_GET, acc_slot);
            }

            // ── Type operations ─────────────────────────────────────────
            ExprKind::IsType {
                expr: inner,
                type_name } => {
                let canon_type = self.canon(type_name);
                if matches!(
                    canon_type.as_str(),
                    "IEnumerable"
                        | "ICollection"
                        | "IList"
                        | "IReadOnlyCollection"
                        | "IReadOnlyList"
                ) {
                    if let ExprKind::Ident(name) = &inner.kind {
                        if self.lookup_var_type_hint(name).is_some_and(|hint| {
                            let bare = hint.split('<').next().unwrap_or(hint).trim();
                            matches!(
                                Self::normalize_type_hint(bare).as_str(),
                                "list" | "arraylist" | "queue" | "stack" | "hashset" | "dictionary"
                            )
                        }) {
                            inst!(self, core_wasm::bool_const, true);
                            return Ok(());
                        }
                    }
                }
                self.compile_expr(inner)?;
                let obj_slot = self.define_local("__is_type_obj");
                self.emit_u16(Op::LOCAL_SET, obj_slot);

                // `IsType` always produces `Value::Bool` — the WASM `REF_IS_*`
                // opcodes return `i32`, so wrap with `emit_i32_to_bool` here.
                // This is language-agnostic: JS normalizes `typeof x === "string"`
                // and `Array.isArray(x)` to `IsType` in the walker; Python uses
                // `isinstance`; VB uses `TypeOf x Is T`. All share the same
                // ECMA-compatible Bool output.
                {
                    let line = self.line;
                    match canon_type.as_str() {
                        "string" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            fn_call!(self, "wasm:js-string", "test", 1);
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "number" | "float" | "double" | "int" | "integer" | "long" | "single"
                        | "decimal" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            fn_call!(self, "wasm:js-number", "test", 1);
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "boolean" | "bool" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            fn_call!(self, "wasm:js-boolean", "test", 1);
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "array" | "list" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            fn_call!(self, "ecma:array", "isArray", 1);
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "function" | "callable" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            inst!(self, recipes::is_func);
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "undefined" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        "object" | "dict" | "map" => {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            inst!(self, recipes::is_object);
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                            return Ok(());
                        }
                        _ => {} // fall through to the general reflection path
                    }
                }

                // `TypeOf x Is Integer` — an Integer value type is a number that
                // is integral. Only the reflection/`TypeOf..Is` path reaches here
                // with these type names (C#'s `is int` desugars via pattern
                // matching), so the type-name check alone is the correct gate.
                if matches!(canon_type.as_str(), "integer" | "int") {
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    fn_call!(self, "wasm:js-number", "test", 1);
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    inst!(self, core_wasm::dup);
                    self.emit(Op::F64_TRUNC);
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    self.chunk().emit_else(line);
                    inst!(self, core_wasm::bool_const, false);
                    self.chunk().emit_end(line);
                    return Ok(());
                }

                let line = self.line;
                let matched_slot = self.define_local("__type_test_matched");
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, matched_slot);

                let canon_type_name = canon_type.clone();
                crate::primitives::reflection::emit_is_instance_of(
                    &mut self.chunks,
                    self.current,
                    obj_slot,
                    &canon_type_name,
                    line,
                );
                self.chunk().emit_if(line);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.chunk().emit_end(line);

                let mut type_name_candidates = vec![canon_type.clone()];
                let short_type = self.canon(&self.reflection_type_short_name(type_name));
                if !type_name_candidates
                    .iter()
                    .any(|candidate| candidate.eq_ignore_ascii_case(&short_type))
                {
                    type_name_candidates.push(short_type);
                }
                let raw_short = type_name
                    .rsplit(['.', '\\'])
                    .next()
                    .unwrap_or(type_name)
                    .trim()
                    .to_string();
                if !type_name_candidates
                    .iter()
                    .any(|candidate| candidate == &raw_short)
                {
                    type_name_candidates.push(raw_short);
                }
                for candidate in type_name_candidates {
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let type_key = self.str_const("__type");
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, type_key);
                    self.emit_const(Value::String(Arc::from(candidate.as_str())));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.chunk().emit_end(line);
                }

                let reflection_matches: Vec<String> = self
                    .reflection_types
                    .keys()
                    .filter(|candidate| {
                        !candidate.eq_ignore_ascii_case(&canon_type)
                            && self.reflection_is_assignable_from(type_name, candidate)
                    })
                    .cloned()
                    .collect();
                for candidate in &reflection_matches {
                    let candidate_name = candidate.clone();
                    crate::primitives::reflection::emit_is_instance_of(
                        &mut self.chunks,
                        self.current,
                        obj_slot,
                        &candidate_name,
                        line,
                    );
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    let type_key = self.str_const("__type");
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, type_key);
                    self.emit_const(Value::String(Arc::from(candidate.as_str())));
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.chunk().emit_end(line);
                }

                if matches!(
                    canon_type.as_str(),
                    "IEnumerable"
                        | "ICollection"
                        | "IList"
                        | "IReadOnlyCollection"
                        | "IReadOnlyList"
                ) {
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    fn_call!(self, "ecma:array", "isArray", 1);
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_ref_type_test(Op::REF_TEST, "list", line);
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.chunk().emit_end(line);

                    for key_name in ["length", "count"] {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        let key = self.str_const(key_name);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, key);
                        self.emit(Op::REF_IS_NULL);
                        self.emit(Op::I32_EQZ);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.chunk().emit_end(line);
                    }

                    if canon_type == "IEnumerable" {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        inst!(self, recipes::is_object);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.chunk().emit_end(line);
                    }
                }

                self.emit_u16(Op::LOCAL_GET, obj_slot);
                let types_key = self.str_const("__types");
                self.emit_struct_field_op(Op::STRUCT_GET, 0, types_key);
                inst!(self, core_wasm::dup);
                self.emit(Op::REF_IS_NULL);
                self.emit(Op::I32_EQZ);
                self.chunk().emit_if(line);
                let types_slot = self.define_local("__type_test_types");
                self.emit_u16(Op::LOCAL_SET, types_slot);
                self.emit_u16(Op::LOCAL_GET, types_slot);
                self.emit_const(Value::String(Arc::from(canon_type.as_str())));
                common::collections::emit_contains(&mut self.chunks, self.current, line);
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.chunk().emit_end(line);
                for candidate in &reflection_matches {
                    self.emit_u16(Op::LOCAL_GET, types_slot);
                    self.emit_const(Value::String(Arc::from(candidate.as_str())));
                    common::collections::emit_contains(&mut self.chunks, self.current, line);
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                    self.chunk().emit_end(line);
                }
                self.chunk().emit_else(line);
                self.emit(Op::DROP);
                self.chunk().emit_end(line);
                // matched_slot holds I32(0) or I32(1); convert to Bool for Go type assertions.
                self.emit_u16(Op::LOCAL_GET, matched_slot);
                self.chunk().emit_if_value(line);
                inst!(self, core_wasm::bool_const, true);
                self.chunk().emit_else(line);
                inst!(self, core_wasm::bool_const, false);
                self.chunk().emit_end(line);
            }

            ExprKind::Cast { expr: inner, .. } => {
                // Pascal allows user functions to shadow builtin type names
                // (`function Double(x: Integer)`). When the parser produced a
                // builtin-style cast node for such a name, honour the user
                // function instead of treating the cast as a no-op.
                if let ExprKind::Cast { type_name, .. } = &expr.kind {
                    let canon_type = self.canon(type_name);
                    if self.profile.namespaces.use_dotnet {
                        if let Some(names) = self.enum_value_names.get(&canon_type) {
                            if let ExprKind::Lit(Literal::Int(n)) = &inner.kind {
                                if let Some(member_name) = names.get(n) {
                                    self.emit_const(Value::String(Arc::from(member_name.as_str())));
                                    return Ok(());
                                }
                            }
                        }
                    }

                    if self.profile.namespaces.use_dotnet && canon_type == "char" {
                        self.compile_expr(inner)?;
                        let num = self.import("ecma:number", "Number");
                        self.emit_host_call(num, 1);
                        self.emit(Op::F64_FLOOR);
                        fn_call!(self, "wasm:js-string", "fromCharCode", 1);
                        return Ok(());
                    }

                    if self.profile.namespaces.use_dotnet {
                        match canon_type.as_str() {
                            "int" | "long" | "short" | "byte" | "uint" | "ulong" | "ushort"
                            | "sbyte" => {
                                let is_char_like =
                                    matches!(&inner.kind, ExprKind::Lit(Literal::Char(_)))
                                        || self.infer_expr_type_hint(inner).is_some_and(|hint| {
                                            Self::normalize_type_hint(&hint) == "char"
                                        });
                                if is_char_like {
                                    // Widening a `char` reads its UTF-16 code
                                    // point — `(int)'A'` is 65. This used to
                                    // push a literal ZERO and stringify it,
                                    // never evaluating the operand at all, so
                                    // every dotnet language answered `0`. The
                                    // conversion belongs to the dotnet adapter
                                    // (`Convert.ToChar` already lives there as
                                    // its inverse), not to this arm.
                                    self.compile_expr(inner)?;
                                    let line = self.line;
                                    self.emit_common("dotnet.char_code", 1, line);
                                    return Ok(());
                                }
                                self.compile_expr(inner)?;
                                let num = self.import("ecma:number", "Number");
                                self.emit_host_call(num, 1);
                                self.emit(Op::F64_TRUNC);
                                return Ok(());
                            }
                            _ => {}
                        }
                    }

                    // Fixed-width integer cast. The spellings and their widths
                    // are profile data (`[integer_cast_widths]`) — the table of
                    // C type names that used to sit here was a per-language
                    // table in shared code.
                    let cast_spelling = canon_type.trim().to_string();
                    if let Some(width) = self.profile.integer_cast_widths.get(&cast_spelling) {
                        let bits = width.bits;
                        let signed = width.signed;
                        self.compile_expr(inner)?;
                        let num = self.import("ecma:number", "Number");
                        self.emit_host_call(num, 1);
                        self.emit(Op::F64_TRUNC);
                        if let Some(bits) = bits {
                            let modulus = 2f64.powi(bits as i32);
                            self.emit_c_unsigned_wrap(modulus);
                            if signed {
                                self.emit_c_signed_wrap_from_unsigned(modulus / 2.0, modulus);
                            }
                        }
                        return Ok(());
                    }
                    if self
                        .profile
                        .float_cast_types
                        .iter()
                        .any(|spelling| spelling == &cast_spelling)
                    {
                        self.compile_expr(inner)?;
                        let num = self.import("ecma:number", "Number");
                        self.emit_host_call(num, 1);
                        return Ok(());
                    }

                    // A cast to an INTEGER type truncates toward zero. The type
                    // names come from the profile's own `[builtin_types] int`
                    // spellings — the hardcoded `"integer" | "int" | "longint"`
                    // list that used to be here was a per-language table in
                    // shared code.
                    if self.profile.integer_cast_truncates {
                        // `canon` folds case per the profile — pascal is
                        // case-insensitive, so `Integer` must match the
                        // lowercase spelling table.
                        if vybe_ast::builtin_types::classify_with(
                            &self.profile.builtin_type_spellings,
                            &self.canon(type_name),
                        ) == Some(vybe_ast::builtin_slots::BuiltinType::Int)
                        {
                            self.compile_expr(inner)?;
                            let line = self.line;
                            common::math::emit_trunc(self.chunk(), line);
                            return Ok(());
                        }
                    }

                    // `TryCast:Target` is a walker-produced cast encoding
                    // (only VB emits it today); gate on the encoding itself, not
                    // the language name.
                    if type_name.contains(':') {
                        if let Some((cast_kind, target_type)) = type_name.split_once(':') {
                            if cast_kind.eq_ignore_ascii_case("TryCast") {
                                let resolved_target = self.resolve_source_type_alias(target_type);
                                let trimmed_target = resolved_target
                                    .trim()
                                    .trim_end_matches('?')
                                    .trim()
                                    .to_string();
                                if self.vb_is_reference_type_hint(&trimmed_target) {
                                    let line = self.line;
                                    self.compile_expr(inner)?;
                                    let value_slot = self.define_local("__vb_trycast_value");
                                    self.emit_u16(Op::LOCAL_SET, value_slot);
                                    let result_slot = self.define_local("__vb_trycast_result");
                                    self.emit_null();
                                    self.emit_u16(Op::LOCAL_SET, result_slot);

                                    self.emit_u16(Op::LOCAL_GET, value_slot);
                                    self.emit(Op::REF_IS_NULL);
                                    self.emit(Op::I32_EQZ);
                                    let non_null_line = self.line;
                                    self.chunk().emit_if(non_null_line);

                                    if Self::is_string_type_hint(&trimmed_target) {
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        fn_call!(self, "wasm:js-string", "test", 1);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.chunk().emit_end(line);
                                    } else if trimmed_target.ends_with("()") {
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        fn_call!(self, "ecma:array", "isArray", 1);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.chunk().emit_end(line);
                                    } else if self.vb_is_object_type_hint(&trimmed_target) {
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        fn_call!(self, "wasm:js-string", "test", 1);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.chunk().emit_end(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        fn_call!(self, "ecma:array", "isArray", 1);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.chunk().emit_end(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        inst!(self, recipes::is_object);
                                        let line = self.line;
                                        self.chunk().emit_if(line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.chunk().emit_end(line);
                                    } else {
                                        let mut expected_names = vec![trimmed_target.clone()];
                                        if let Some(class_name) = self
                                            .resolve_pending_class_name_for_type_hint(
                                                &trimmed_target,
                                            )
                                        {
                                            if !expected_names.iter().any(|candidate| {
                                                candidate.eq_ignore_ascii_case(&class_name)
                                            }) {
                                                expected_names.push(class_name.clone());
                                            }
                                            let short_name = class_name
                                                .rsplit('.')
                                                .next()
                                                .unwrap_or(class_name.as_str())
                                                .to_string();
                                            if !expected_names.iter().any(|candidate| {
                                                candidate.eq_ignore_ascii_case(&short_name)
                                            }) {
                                                expected_names.push(short_name);
                                            }
                                        } else if let Some(short_name) =
                                            trimmed_target.rsplit('.').next()
                                        {
                                            if !expected_names.iter().any(|candidate| {
                                                candidate.eq_ignore_ascii_case(short_name)
                                            }) {
                                                expected_names.push(short_name.to_string());
                                            }
                                        }

                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        inst!(self, recipes::is_object);
                                        let object_line = self.line;
                                        self.chunk().emit_if(object_line);
                                        let match_slot = self.define_local("__vb_trycast_match");
                                        self.emit_const(Value::I32(0));
                                        self.emit_u16(Op::LOCAL_SET, match_slot);
                                        for expected in &expected_names {
                                            self.emit_u16(Op::LOCAL_GET, value_slot);
                                            let type_key = self.str_const("__type");
                                            self.emit_struct_field_op(Op::STRUCT_GET, 0, type_key);
                                            self.emit_const(Value::String(Arc::from(
                                                expected.as_str(),
                                            )));
                                            {
                                                let line = self.line;
                                                crate::primitives::ops::emit_dyn_eq(
                                                    self.chunk(),
                                                    line,
                                                );
                                            };
                                            crate::primitives::ops::emit_dyn_to_bool(
                                                self.chunk(),
                                                object_line,
                                            );
                                            self.chunk().emit_if(object_line);
                                            self.emit_const(Value::I32(1));
                                            self.emit_u16(Op::LOCAL_SET, match_slot);
                                            self.chunk().emit_end(object_line);
                                        }
                                        for expected in &expected_names {
                                            self.emit_u16(Op::LOCAL_GET, value_slot);
                                            let types_key = self.str_const("__types");
                                            self.emit_struct_field_op(Op::STRUCT_GET, 0, types_key);
                                            inst!(self, core_wasm::dup);
                                            self.emit(Op::REF_IS_NULL);
                                            self.emit(Op::I32_EQZ);
                                            self.chunk().emit_if(object_line);
                                            let types_slot =
                                                self.define_local("__vb_trycast_types");
                                            self.emit_u16(Op::LOCAL_SET, types_slot);
                                            self.emit_u16(Op::LOCAL_GET, types_slot);
                                            self.emit_const(Value::String(Arc::from(
                                                expected.as_str(),
                                            )));
                                            common::collections::emit_contains(
                                                &mut self.chunks,
                                                self.current,
                                                line,
                                            );
                                            crate::primitives::ops::emit_dyn_to_bool(
                                                self.chunk(),
                                                object_line,
                                            );
                                            self.chunk().emit_if(object_line);
                                            self.emit_const(Value::I32(1));
                                            self.emit_u16(Op::LOCAL_SET, match_slot);
                                            self.chunk().emit_end(object_line);
                                            self.chunk().emit_else(object_line);
                                            self.emit(Op::DROP);
                                            self.chunk().emit_end(object_line);
                                        }
                                        self.emit_u16(Op::LOCAL_GET, match_slot);
                                        self.chunk().emit_if(object_line);
                                        self.emit_u16(Op::LOCAL_GET, value_slot);
                                        self.emit_u16(Op::LOCAL_SET, result_slot);
                                        self.chunk().emit_end(object_line);
                                        self.chunk().emit_end(object_line);
                                    }

                                    self.chunk().emit_end(non_null_line);
                                    self.emit_u16(Op::LOCAL_GET, result_slot);
                                    return Ok(());
                                }
                            }
                        }
                    }

                    if let Some(user_type) = self.user_value_type_name_from_hint(type_name) {
                        if matches!(&inner.kind, ExprKind::Object(_)) {
                            let ctor_global = {
                                let overload =
                                    crate::primitives::classes::ctor_global_for(&user_type, 0);
                                if self.defined_globals.contains(&overload) {
                                    overload
                                } else {
                                    user_type.clone()
                                }
                            };

                            let source_slot = self.define_local("__cast_struct_source");
                            self.compile_expr(inner)?;
                            self.emit_u16(Op::LOCAL_SET, source_slot);

                            let value_slot = self.define_local("__cast_struct_value");
                            self.emit_global_read(&ctor_global);
                            self.emit_u8(Op::CALL_REF, 0);
                            self.emit_u16(Op::LOCAL_SET, value_slot);

                            if let Some(fields) = self
                                .pending_classes
                                .get(&user_type)
                                .map(|pending| pending.fields.clone())
                            {
                                for field_name in fields {
                                    let member_slot = self.define_local("__cast_struct_member");
                                    let field_idx = self.str_const(&field_name);
                                    self.emit_u16(Op::LOCAL_GET, source_slot);
                                    self.emit_struct_field_op(Op::STRUCT_GET, 0, field_idx);
                                    self.emit_u16(Op::LOCAL_SET, member_slot);

                                    self.emit_u16(Op::LOCAL_GET, member_slot);
                                    fn_call!(self, "wasm:js-undefined", "test", 1);
                                    self.emit(Op::I32_EQZ);
                                    let set_line = self.line;
                                    self.chunk().emit_if(set_line);
                                    self.emit_u16(Op::LOCAL_GET, value_slot);
                                    self.emit_u16(Op::LOCAL_GET, member_slot);
                                    self.emit_struct_field_op(Op::STRUCT_SET, 0, field_idx);
                                    self.chunk().emit_end(set_line);
                                }
                            }

                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            return Ok(());
                        }
                    }

                    let shadows_cast = self.defined_functions.contains(&canon_type)
                        || (!self.case_sensitive
                            && self
                                .defined_functions
                                .iter()
                                .any(|name| name.eq_ignore_ascii_case(type_name)));
                    if shadows_cast {
                        self.emit_var_get(type_name);
                        self.compile_expr(inner)?;
                        self.emit_u8(Op::CALL_REF, 1);
                        return Ok(());
                    }
                }

                // Cast is otherwise a no-op in our dynamic VM.
                self.compile_expr(inner)?;
            }

            ExprKind::DefaultOf(type_name) => {
                let normalized = Self::normalize_type_hint(type_name);
                match normalized.as_str() {
                    "int" | "long" | "short" | "byte" | "uint" | "ulong" | "ushort" | "sbyte"
                    | "double" | "float" | "decimal" => {
                        self.emit_const(Value::I32(0));
                    }
                    "bool" => {
                        self.emit_const(Value::Bool(false));
                    }
                    "char" => {
                        self.emit_const(Value::String(Arc::from("\0")));
                    }
                    _ => {
                        self.emit_null();
                    }
                }
            }

            ExprKind::TypeOf(inner) => {
                if self.profile.ecma_typeof_operator {
                    if let ExprKind::Ident(name) = &inner.kind {
                        let is_callable_global = !self.has_accessible_local_binding(name)
                            && (matches!(
                                name.as_str(),
                                "eval"
                                    | "parseInt"
                                    | "parseFloat"
                                    | "Function"
                                    | "Object"
                                    | "Boolean"
                                    | "Number"
                                    | "String"
                                    | "Array"
                                    | "Symbol"
                                    | "BigInt"
                                    | "Date"
                                    | "RegExp"
                                    | "Promise"
                                    | "Proxy"
                                    | "Map"
                                    | "Set"
                                    | "WeakMap"
                                    | "WeakSet"
                                    | "ArrayBuffer"
                                    | "SharedArrayBuffer"
                                    | "DataView"
                                    | "Int8Array"
                                    | "Uint8Array"
                                    | "Uint8ClampedArray"
                                    | "Int16Array"
                                    | "Uint16Array"
                                    | "Int32Array"
                                    | "Uint32Array"
                                    | "Float32Array"
                                    | "Float64Array"
                                    | "BigInt64Array"
                                    | "BigUint64Array"
                            ) || self.profile.lookup_builtin(name).is_some());
                        if is_callable_global {
                            self.emit_const(Value::String(Arc::from("function")));
                            return Ok(());
                        }
                    }
                }
                let saved_typeof = self.in_typeof_operand;
                self.in_typeof_operand = true;
                let inner_res = self.compile_expr(inner);
                self.in_typeof_operand = saved_typeof;
                inner_res?;
                if self.profile.ecma_typeof_operator {
                    // ECMA-262 §13.5.3 Table 41: arrays are "object",
                    // not "array". The VM's REF_TYPEOF emits "array"
                    // (Vybe-specific), so JS routes through the host
                    // helper that returns spec-compliant tags.
                    let idx = self.import("ecma:value", "typeof");
                    self.emit_host_call(idx, 1);
                } else {
                    fn_call!(self, "ecma:value", "typeof", 1);
                }
            }

            // ── NullCoalesce ────────────────────────────────────────────
            ExprKind::NullCoalesce { left, right } => {
                self.compile_expr(left)?;
                inst!(self, core_wasm::dup);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.emit(Op::DROP);
                self.compile_expr(right)?;
                self.chunk().emit_else(line);
                self.chunk().emit_end(line);
            }

            // ── Spread ──────────────────────────────────────────────────
            ExprKind::Spread(inner) => {
                self.compile_expr(inner)?;
                // SPREAD is array-only in the VM (matches WASM
                // `array.copy_into` semantics). Iterables that aren't
                // arrays — Set, Map, String — get coerced to an array
                // first via the polymorphic Symbol.iterator helper
                // (`ecma:object.iterForOf`). Generators (Continuation
                // values) need spec WASM stack-switching `resume`
                // (emit_next) to drive their iterator protocol, which a
                // host fn can't do — route them through the
                // `__stdlib_drain_generator` bytecode helper.
                common::collections::emit_spread_iterable(
                    &mut self.chunks,
                    self.current,
                    self.line,
                );
            }

            // ── Await ───────────────────────────────────────────────────
            // The async model — one vocabulary for every language's
            // spelling, one lowering (`primitives/async_ops.rs`).
            ExprKind::Async(op) => {
                self.emit_async(op)?;
            }
            ExprKind::Chan(op) => {
                self.emit_chan(op)?;
            }
            ExprKind::Await(inner) => {
                // ECMA-262 §27.2: WASM JSPI suspend point, lowered to the spec
                // stack-switching `suspend` (AWAIT_SUSPEND_TAG). The VM unwraps
                // fulfilled, throws rejected, suspends the fiber on pending, and
                // passes non-promise values through unchanged.
                self.compile_expr(inner)?;
                let line = self.line;
                crate::primitives::functions::emit_await(self.chunk(), line);
            }

            // ── Yield ───────────────────────────────────────────────────
            ExprKind::Yield(val) => {
                if let Some(v) = val {
                    if let Some((key_expr, value_expr)) = self.generator_keyed_yield_parts(v) {
                        self.compile_expr(key_expr)?;
                        let key_slot = self.define_local("__yield_key");
                        self.emit_u16(Op::LOCAL_SET, key_slot);

                        self.compile_expr(value_expr)?;
                        let payload_value_slot = self.define_local("__yield_payload_value");
                        self.emit_u16(Op::LOCAL_SET, payload_value_slot);

                        self.emit_next_generator_payload_id();
                        let payload_id_slot = self.define_local("__yield_payload_id");
                        self.emit_u16(Op::LOCAL_SET, payload_id_slot);

                        self.emit_generator_payload_store();
                        self.emit_u16(Op::LOCAL_GET, payload_id_slot);
                        self.emit_u16(Op::LOCAL_GET, payload_value_slot);
                        let line = self.line;
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);

                        common::dict::emit_new(&mut self.chunks, self.current, line);
                        inst!(self, core_wasm::dup);
                        self.emit_const(Value::Bool(true));
                        let marker_key = self.str_const("__vybe_generator_yield");
                        self.emit_struct_field_op(Op::STRUCT_SET, 0, marker_key);
                        inst!(self, core_wasm::dup);
                        self.emit_u16(Op::LOCAL_GET, key_slot);
                        let key_key = self.str_const("key");
                        self.emit_struct_field_op(Op::STRUCT_SET, 0, key_key);
                        inst!(self, core_wasm::dup);
                        self.emit_u16(Op::LOCAL_GET, payload_id_slot);
                        let payload_id_key = self.str_const("payload_id");
                        self.emit_struct_field_op(Op::STRUCT_SET, 0, payload_id_key);
                    } else {
                        self.compile_expr(v)?;
                    }
                } else {
                    self.emit_null();
                }
                let line = self.line;
                crate::primitives::generators::emit_suspend(self.chunk(), line);
                self.emit_generator_resume_value()?;
            }

            ExprKind::YieldFrom(inner) => {
                // ECMA-262 §15.5 `yield*`: drain the inner iterable,
                // re-yielding each value through the enclosing
                // generator. Continuations use WASM stack-switching
                // `GEN_NEXT`; plain iterables are first drained through
                // the shared collection surface and yielded value-by-value.
                self.compile_expr(inner)?;
                let gen_slot = self.define_local("__yield_star_gen");
                let val_slot = self.define_local("__yield_star_val");
                let has_more_slot = self.define_local("__yield_star_has_more");
                let result_slot = self.define_local("__yield_star_result");
                self.emit_u16(Op::LOCAL_SET, gen_slot);

                self.emit_u16(Op::LOCAL_GET, gen_slot);
                let is_gen_idx = self.import("ecma:value", "isGenerator");
                self.emit_host_call(is_gen_idx, 1);
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);

                let gen_block = self.chunk().emit_block(line);
                let (gen_loop, _) = self.chunk().emit_loop_s(line);
                self.emit_u16(Op::LOCAL_GET, gen_slot);
                let line = self.line;
                crate::primitives::generators::emit_next(self.chunk(), line);
                // After GEN_NEXT: stack top is has_more (i32), under it value.
                self.emit_u16(Op::LOCAL_SET, has_more_slot);
                self.emit_u16(Op::LOCAL_SET, val_slot);
                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.emit(Op::I32_EQZ);
                self.chunk().emit_br_if(1, line);
                self.emit_u16(Op::LOCAL_GET, val_slot);
                let line = self.line;
                crate::primitives::generators::emit_suspend(self.chunk(), line);
                self.emit(Op::DROP);
                self.chunk().emit_br(0, line);
                self.chunk().emit_end(line);
                self.chunk().patch_loop(gen_loop);
                self.chunk().emit_end(line);
                self.chunk().patch_block(gen_block);
                // `yield from` / `yield*` evaluates to the delegated
                // generator's completion value. `GEN_NEXT` leaves that
                // final value in `val_slot` on the terminating step.
                self.emit_u16(Op::LOCAL_GET, val_slot);
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, gen_slot);
                common::collections::emit_iter_for_of(&mut self.chunks, self.current, self.line);
                let iter_slot = self.define_local("__yield_star_iter");
                let idx_slot = self.define_local("__yield_star_idx");
                let len_slot = self.define_local("__yield_star_len");
                self.emit_u16(Op::LOCAL_SET, iter_slot);

                self.emit_const(Value::F64(0.0));
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, iter_slot);
                let line = self.line;
                common::collections::emit_array_length(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, len_slot);

                let iter_block = self.chunk().emit_block(line);
                let (iter_loop, _) = self.chunk().emit_loop_s(line);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, len_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
                self.chunk().emit_br_if(1, line);
                self.emit_u16(Op::LOCAL_GET, iter_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::ARRAY_GET);
                let line = self.line;
                crate::primitives::generators::emit_suspend(self.chunk(), line);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_const(Value::F64(1.0));
                self.emit(Op::F64_ADD);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.chunk().emit_br(0, line);
                self.chunk().emit_end(line);
                self.chunk().patch_loop(iter_loop);
                self.chunk().emit_end(line);
                self.chunk().patch_block(iter_block);
                self.emit_null();
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_GET, result_slot);
            }

            // ── AddressOf (VB) ──────────────────────────────────────────
            ExprKind::AddressOf(name) => {
                let parts: Vec<&str> = name.split('.').filter(|part| !part.is_empty()).collect();
                if parts.is_empty() {
                    self.emit_null();
                    return Ok(());
                }

                let self_kw = self.profile.self_keyword.clone();
                let is_self_qualified = parts
                    .first()
                    .map(|part| {
                        if self.case_sensitive {
                            *part == self_kw || *part == "Me"
                        } else {
                            part.eq_ignore_ascii_case(&self_kw) || part.eq_ignore_ascii_case("Me")
                        }
                    })
                    .unwrap_or(false);

                let bound_parts: Option<&[&str]> = if self.current_class.is_some() {
                    if is_self_qualified && parts.len() > 1 {
                        Some(&parts[1..])
                    } else if parts.len() == 1
                        && self.defined_class_methods.contains(&self.canon(parts[0]))
                    {
                        Some(&parts[..])
                    } else {
                        None
                    }
                } else {
                    None
                };

                if let Some(method_parts) = bound_parts {
                    if self.emit_self_ref() {
                        for part in method_parts {
                            let idx = self.str_const(&self.canon(part));
                            self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                        }
                        return Ok(());
                    }
                }

                self.emit_var_get(parts[0]);
                for part in &parts[1..] {
                    let idx = self.str_const(&self.canon(part));
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                }
            }

            // ── SuperCall (VB/Python) ───────────────────────────────────
            ExprKind::SuperCall { method, args } => {
                let self_kw = self.profile.self_keyword.clone();
                let ctor_name = self.profile.constructor_name.clone();
                let is_ctor_call = method.is_none()
                    || method.as_ref().map_or(false, |m| {
                        if self.case_sensitive {
                            m == &ctor_name || m == "new" || m == "__init__"
                        } else {
                            m.eq_ignore_ascii_case(&ctor_name)
                                || m.eq_ignore_ascii_case("new")
                                || m.eq_ignore_ascii_case("__init__")
                        }
                    });

                if is_ctor_call {
                    // super() / MyBase.New(args) → call parent constructor
                    if let Some(ref class_name) = self.current_class.clone() {
                        if let Some(parent_name) = self
                            .pending_classes
                            .get(class_name.as_str())
                            .and_then(|c| c.parent.clone())
                        {
                            if !self.shadows_builtin_type(&parent_name)
                                && common::errors::is_exception_type(&parent_name)
                            {
                                let arg_exprs: Vec<&Expression> =
                                    args.iter().map(|arg| &arg.value).collect();
                                self.emit_js_exception_ctor_value(&parent_name, &arg_exprs)?;
                                if let Some(slot) = self
                                    .scope()
                                    .resolve(&self_kw)
                                {
                                    inst!(self, core_wasm::dup);
                                    self.emit_u16(Op::LOCAL_SET, slot);
                                }
                                return Ok(());
                            }
                            // §13.3.7.2 (JS): super() may only run once —
                            // a second call sees this_slot already
                            // initialized and throws a ReferenceError.
                            if let Some((ctx_chunk, ctx_slot)) = self.js_derived_ctor_ctx {
                                if ctx_chunk == self.current {
                                    let l = self.line;
                                    crate::primitives::classes::emit_super_once_guard(
                                        self.chunk(),
                                        ctx_slot,
                                        l,
                                    );
                                }
                            }
                            // Framework GUI control parent (`MyBase.New()` /
                            // `super()` over `Form`/`Button`/…): construct via
                            // the `vybe:gui` host factory, the same GUI-direct
                            // path the auto-base construction uses. Control
                            // leaves no longer have a ctor global to CALL_REF.
                            if self.is_framework_control_parent(&parent_name) {
                                let canonical = common::gui::canonical_control_name(&parent_name);
                                for a in args {
                                    self.compile_expr(&a.value)?;
                                }
                                let line = self.line;
                                self.emit_control_element(&parent_name, args.len() as u8, line);
                            } else {
                                self.emit_var_get(&parent_name);
                                for a in args {
                                    self.compile_expr(&a.value)?;
                                }
                                self.emit_u8(Op::CALL_REF, args.len() as u8);
                            }
                            if let Some(slot) = self
                                .scope()
                                .resolve(&self_kw)
                            {
                                inst!(self, core_wasm::dup);
                                self.emit_u16(Op::LOCAL_SET, slot);
                            }
                        } else {
                            self.emit_null();
                        }
                    } else {
                        self.emit_null();
                    }
                } else if let Some(mname) = method {
                    let parent_name = self
                        .current_class
                        .as_ref()
                        .and_then(|class_name| self.pending_classes.get(class_name.as_str()))
                        .and_then(|pc| pc.parent.clone());
                    if let Some(parent_name) = parent_name {
                        let parent_canon = self.canon(&parent_name);
                        let method_name = self.canon(mname);
                        let method_idx = self.str_const(&method_name);
                        self.emit_var_get(&parent_canon);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, method_idx);

                        if self.profile.ambient_this_binding {
                            let saved_js_this = self.save_js_this("__js_prev_this_super_expr");
                            if let Some(self_slot) = self
                                .scope()
                                .resolve(&self_kw)
                            {
                                self.emit_u16(Op::LOCAL_GET, self_slot);
                            } else {
                                self.emit_global_read("__js_this");
                            }
                            self.set_js_this_from_stack();
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            let result_slot = self.define_local("__js_super_expr_result");
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.restore_js_this(saved_js_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else if let Some(self_slot) = self
                            .scope()
                            .resolve(&self_kw)
                        {
                            self.emit_u16(Op::LOCAL_GET, self_slot);
                            for a in args {
                                self.compile_expr(&a.value)?;
                            }
                            self.emit_u8(Op::CALL_REF, (args.len() + 1) as u8);
                        } else {
                            self.emit_null();
                        }
                    } else {
                        self.emit_null();
                    }
                } else {
                    self.emit_null();
                }
            }

            // ── Comprehension (Python) ──────────────────────────────────
            ExprKind::Comprehension {
                kind,
                element,
                generators } => {
                use crate::ast::ComprehensionKind;
                let line = self.line;
                let is_dict = *kind == ComprehensionKind::Dict;
                let is_set = *kind == ComprehensionKind::Set;
                // Dict comprehension builds a Map (same as a dict literal) so
                // non-string keys keep their type — Python dict === PHP array.
                //
                // `ComprehensionKind::Dict` already SAYS this is a dict; the
                // profile flag it used to be ANDed with was a per-language veto
                // over information the AST had already declared. `is_dict`
                // alone is the condition, and the `else if is_dict` arm below
                // is now dead — kept only until a language declares a dict
                // comprehension that genuinely wants object semantics.
                let dict_as_map = is_dict;

                // Build the accumulator: dict → Map/Object, set/list/gen → Array
                if dict_as_map {
                    common::collections::emit_map_new(&mut self.chunks, self.current, line);
                } else if is_dict {
                    common::dict::emit_new(&mut self.chunks, self.current, line);
                } else {
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                }
                let result_slot = self.define_local("__comp_result");
                self.emit_u16(Op::LOCAL_SET, result_slot);

                use crate::primitives::loops::LoopState;
                // Compile each generator (nested for-clauses)
                let mut cond_blocks = 0usize;
                let mut loop_info: Vec<(u16, LoopState)> = Vec::new();
                for generator in generators.iter() {
                    self.compile_expr(&generator.iter)?;
                    // Materialize the source so the index loop below can walk it:
                    // a lazy generator (e.g. `range()` / a genexpr) drains via
                    // stack-switching; other iterables use the natural per-type
                    // iteration (dict → keys for Python) when the profile opts in.
                    {
                        let line = self.line;
                        let src_slot = self.define_local("__comp_src");
                        self.emit_u16(Op::LOCAL_SET, src_slot);
                        self.emit_u16(Op::LOCAL_GET, src_slot);
                        let is_gen = self.import("ecma:value", "isGenerator");
                        self.emit_host_call(is_gen, 1);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if_value(line);
                        self.emit_u16(Op::LOCAL_GET, src_slot);
                        common::generators::emit_drain_into_array(
                            &mut self.chunks,
                            self.current,
                            line,
                        );
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, src_slot);
                        if self.profile.for_in_object_yields_keys {
                            common::collections::emit_iter_natural(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                        }
                        self.chunk().emit_end(line);
                    }
                    let arr_slot = self.define_local("__comp_iter");
                    self.emit_u16(Op::LOCAL_SET, arr_slot);
                    let idx_slot = self.define_local("__comp_idx");
                    let lp = common::loops::emit_for_in_start(
                        &mut self.chunks,
                        self.current,
                        arr_slot,
                        idx_slot,
                        line,
                    );
                    // Tuple target (`for a, b in …`): store the element, then bind
                    // each component from its position. Single-Ident targets keep
                    // the original single-slot binding; only the previously
                    // unbindable non-Ident path gains destructuring.
                    let tuple_names: Option<Vec<String>> = match &generator.target.kind {
                        ExprKind::Tuple(parts) => Some(
                            parts
                                .iter()
                                .map(|p| match &p.kind {
                                    ExprKind::Ident(n) => n.clone(),
                                    _ => "_".to_string() })
                                .collect(),
                        ),
                        _ => None };
                    if let Some(names) = tuple_names {
                        let elem_slot = self.define_local("__comp_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        for (k, nm) in names.iter().enumerate() {
                            let slot = self.define_local(nm);
                            self.emit_u16(Op::LOCAL_GET, elem_slot);
                            self.emit_const(Value::I32(k as i32));
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                            self.emit_u16(Op::LOCAL_SET, slot);
                        }
                    } else {
                        let var_name = match &generator.target.kind {
                            ExprKind::Ident(n) => n.clone(),
                            _ => "__comp_var".to_string() };
                        let var_slot = self.define_local(&var_name);
                        self.emit_u16(Op::LOCAL_SET, var_slot);
                    }

                    for cond_expr in &generator.conditions {
                        self.compile_expr(cond_expr)?;
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        self.chunk().emit_if(line);
                        cond_blocks += 1;
                    }
                    loop_info.push((idx_slot, lp));
                }

                // Emit the body
                if is_dict {
                    // element is Array([key, val]); store into dict and track __keys
                    if let ExprKind::Array(kv) = &element.kind {
                        if kv.len() == 2 {
                            let key_expr = &kv[0].value;
                            let val_expr = &kv[1].value;
                            // Compile key and save a copy for __keys tracking
                            self.compile_expr(key_expr)?;
                            let key_slot = self.define_local("__comp_key");
                            self.emit_u16(Op::LOCAL_SET, key_slot);
                            // [dict, key, val] → ARRAY_SET → drops from stack
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            self.emit_u16(Op::LOCAL_GET, key_slot);
                            self.compile_expr(val_expr)?;
                            let l = self.line;
                            common::collections::emit_set(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP);
                            // Ordinary objects need explicit __keys order tracking;
                            // a Map tracks insertion order internally, so skip it.
                            if !dict_as_map {
                                self.emit_u16(Op::LOCAL_GET, result_slot);
                                let keys_key = self.str_const("__keys");
                                self.emit_struct_field_op(Op::STRUCT_GET, 0, keys_key);
                                self.emit_u16(Op::LOCAL_GET, key_slot);
                                let l = self.line;
                                common::collections::emit_push(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP);
                            }
                        }
                    }
                } else {
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.compile_expr(element)?;
                    let l = self.line;
                    common::collections::emit_push(&mut self.chunks, self.current, l);
                    self.emit(Op::DROP);
                }

                for _ in 0..cond_blocks {
                    self.chunk().emit_end(line);
                }
                // Close loops in reverse order
                for (idx_slot, lp) in loop_info.into_iter().rev() {
                    common::loops::emit_for_in_end(
                        &mut self.chunks,
                        self.current,
                        idx_slot,
                        lp,
                        line,
                    );
                }

                self.emit_u16(Op::LOCAL_GET, result_slot);

                // Set comprehension: convert Array → Set via ecma:set.fromIterable
                if is_set {
                    let idx = self.import("ecma:set", "fromIterable");
                    self.emit_host_call(idx, 1);
                }
            }

            // ── Slice (Python) ──────────────────────────────────────────
            ExprKind::Slice { lower, upper, step } => {
                // Stack on entry (from Index parent): [obj]
                let line = self.line;
                if step.is_none() {
                    let obj_slot = self.define_local("__py_slice_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    if let Some(l) = lower {
                        self.compile_expr(l)?;
                    } else {
                        inst!(self, core_wasm::i32_const, 0);
                    }
                    if let Some(u) = upper {
                        self.compile_expr(u)?;
                    } else {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        common::collections::emit_len(&mut self.chunks, self.current, line);
                    }
                    // Contiguous slice via ecma:array.slice (string→substring /
                    // array→array, negative-wrap + clamp). Home: crate::primitives::slices.
                    crate::primitives::slices::emit_contiguous(
                        &mut self.chunks,
                        self.current,
                        line,
                    );
                } else {
                    // Strided slice → [obj, lower, upper, step] (NULL = absent);
                    // obj is already on the stack from the Index parent.
                    // crate::primitives::slices emits the CPython normalization +
                    // strided copy inline; the step==0 quirk comes from the
                    // `slice_step_zero_raises` profile property.
                    if let Some(l) = lower {
                        self.compile_expr(l)?;
                    } else {
                        self.emit_null();
                    }
                    if let Some(u) = upper {
                        self.compile_expr(u)?;
                    } else {
                        self.emit_null();
                    }
                    if let Some(s) = step {
                        self.compile_expr(s)?;
                    } else {
                        self.emit_null();
                    }
                    let opts = crate::primitives::slices::Options::new(
                        self.profile.slice_step_zero_raises,
                    );
                    crate::primitives::slices::emit_stepped(
                        &mut self.chunks,
                        self.current,
                        line,
                        opts,
                    );
                }
            }

            // ── Walrus (Python :=) ──────────────────────────────────────
            ExprKind::Walrus { target, value } => {
                self.compile_expr(value)?;
                inst!(self, core_wasm::dup);
                self.compile_assign_target(target)?;
            }

            // ── Void (JS) ───────────────────────────────────────────────
            ExprKind::Void(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::DROP);
                inst!(self, core_wasm::undefined); // ECMA-262 §13.5.2: void → undefined
            }

            // ── Delete (JS expression) ──────────────────────────────────
            ExprKind::Delete(inner) => {
                // delete obj.prop → ecma:object.delete(obj, key), returns true.
                // Proxy modules route through ecma:proxy.deleteProperty so the
                // deleteProperty trap fires (non-proxy targets fall through).
                if matches!(
                    inner.as_ref(),
                    Expression {
                        kind: ExprKind::Member { object, .. },
                        ..
                    } if matches!(object.kind, ExprKind::Super)
                ) || matches!(
                    inner.as_ref(),
                    Expression {
                        kind: ExprKind::Index { object, .. },
                        ..
                    } if matches!(object.kind, ExprKind::Super)
                ) {
                    self.emit_const(Value::String(Arc::from("Cannot delete super property")));
                    let line = self.line;
                    self.emit_js_exception_ctor_from_message_value("ReferenceError")?;
                    common::errors::emit_throw(self.chunk(), line);
                    return Ok(());
                }
                let delete_import: (&str, &str) = if self.uses_proxy {
                    ("ecma:proxy", "deleteProperty")
                } else {
                    ("ecma:object", "delete")
                };
                if let ExprKind::Member { object, field, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    let idx = self.import(delete_import.0, delete_import.1);
                    self.emit_host_call(idx, 2);
                    if self.in_strict {
                        let result_slot = self.define_local("__strict_delete_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.emit(Op::I32_EQZ);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::String(Arc::from(
                            "Cannot delete non-configurable property",
                        )));
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        common::errors::emit_throw(self.chunk(), line);
                        self.chunk().emit_end(line);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                } else if let ExprKind::Index { object, index, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    let idx = self.import(delete_import.0, delete_import.1);
                    self.emit_host_call(idx, 2);
                    if self.in_strict {
                        let result_slot = self.define_local("__strict_delete_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.emit(Op::I32_EQZ);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::String(Arc::from(
                            "Cannot delete non-configurable property",
                        )));
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        common::errors::emit_throw(self.chunk(), line);
                        self.chunk().emit_end(line);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                } else {
                    self.compile_expr(inner)?;
                    self.emit(Op::DROP);
                    inst!(self, core_wasm::bool_const, true);
                }
            }

            // ── Destructure (JS) ────────────────────────────────────────
            ExprKind::Destructure(_) => {
                // Destructure patterns are handled at assignment/declaration sites
                self.emit_null();
            }

            // ── Sequence (JS comma operator) ────────────────────────────
            ExprKind::Sequence(exprs) => {
                for (i, e) in exprs.iter().enumerate() {
                    self.compile_expr(e)?;
                    if i < exprs.len() - 1 {
                        self.emit(Op::DROP);
                    }
                }
            }

            // ── ClassExpr (JS) ──────────────────────────────────────────
            ExprKind::ClassExpr {
                name,
                parent,
                interfaces,
                members } => {
                let class_name = name
                    .clone()
                    .unwrap_or_else(|| format!("__anonymous_class_{}", self.chunks.len()));
                let class_name = self.canon(&class_name);
                let parent_name: Option<String> = if let Some(p) = parent {
                    // A base named by a plain identifier that is already a
                    // defined class is known STATICALLY — use that name. The
                    // synthetic-global indirection below exists for the general
                    // `extends <arbitrary expression>` form, but routing a
                    // static base through it hands `register_type` the name
                    // `__extends_<class>_<n>`, which is not a registered type;
                    // `load_type_table` resolves it to `None` and the class
                    // lands in the registry with NO parent. Invisible while the
                    // rtt is 0 (the `__types` string scan still answered), fatal
                    // once a real `type_id` is stamped: `$o instanceof Base`
                    // false for `new class extends Base{}`.
                    let static_base = match &p.kind {
                        ExprKind::Ident(n) => {
                            let c = self.canon(n);
                            self.defined_classes.contains(&c).then_some(c)
                        }
                        _ => None };
                    if let Some(base) = static_base {
                        Some(base)
                    } else {
                        let synth_parent =
                            self.canon(&format!("__extends_{}_{}", class_name, self.chunks.len()));
                        self.defined_globals.insert(synth_parent.clone());
                        self.compile_expr(p)?;
                        self.emit_global_write(&synth_parent);
                        Some(synth_parent)
                    }
                } else {
                    None
                };
                self.defined_globals.insert(class_name.clone());
                // Register as a defined class too: an anonymous class read back
                // by name inside a PHP function must be recognised as a class,
                // otherwise emit_var_get's in-function global guard emits NULL
                // (PHP globals aren't visible in function scope) and the `new`
                // constructs null. Class names are language-agnostic here.
                self.defined_classes.insert(class_name.clone());
                let parents: Vec<String> = parent_name.into_iter().collect();
                let saved_expr_js_this = if self.profile.ambient_this_binding {
                    Some(self.save_js_this(&format!("__js_prev_this_class_expr_{}", class_name)))
                } else {
                    None
                };
                crate::primitives::class_normalize::emit::emit_class_from_ast(
                    self,
                    expr.span.clone(),
                    &class_name,
                    &parents,
                    interfaces,
                    members,
                    &crate::ast::ClassModifiers::default(),
                    vybe_ast::ValueSemantics::default(),
                )?;
                if let Some(saved) = saved_expr_js_this {
                    self.restore_js_this(saved);
                }
                self.emit_global_read(&class_name);
            }

            // ── FunctionExpr (JS) ───────────────────────────────────────
            ExprKind::FunctionExpr(stmt) => {
                if let StmtKind::FunctionDecl {
                    name,
                    params,
                    return_type,
                    body,
                    is_sub,
                    is_generator,
                    handles,
                    is_async,
                    ..
                } = &stmt.kind
                {
                    let fn_name = if name.is_empty() {
                        format!("__anon_fn_{}", self.chunks.len())
                    } else {
                        name.clone()
                    };
                    self.compile_function_decl(
                        &fn_name,
                        params,
                        return_type,
                        body,
                        *is_sub,
                        *is_generator,
                        handles,
                        *is_async,
                    )?;
                    self.emit_var_get(&fn_name);
                } else {
                    self.emit_null();
                }
            }

            // ── Range ───────────────────────────────────────────────────
            ExprKind::Range {
                start,
                end,
                inclusive } => {
                self.compile_expr(start)?;
                self.compile_expr(end)?;
                let line = self.line;
                common::collections::emit_range(
                    &mut self.chunks,
                    self.current,
                    2,
                    *inclusive,
                    line,
                );
            }

            // ── StaticAccess (PHP) ──────────────────────────────────────
            ExprKind::StaticAccess { class, member } => {
                if let (ExprKind::Ident(class_name), ExprKind::Ident(member_name)) =
                    (&class.kind, &member.kind)
                {
                    if self.private_member_access_forbidden(member_name) {
                        self.emit_private_access_denied(member_name)?;
                        return Ok(());
                    }
                    if let Some(value) = self.enum_member_ordinal(class_name, member_name) {
                        self.emit_const(Value::F64(value as f64));
                        return Ok(());
                    }

                    let compound = self.canon(&format!("{}.{}", class_name, member_name));
                    if let Some(cv) = self.profile.lookup_constant(&compound) {
                        match cv {
                            ConstantValue::Bool(b) => self.emit_const(Value::Bool(*b)),
                            ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                            ConstantValue::Str(s) => {
                                self.emit_const(Value::String(Arc::from(s.as_str())))
                            }
                        }
                        return Ok(());
                    }
                    // User-defined class constant — resolve as a global
                    if self.defined_globals.contains(&compound) {
                        self.emit_global_read(&compound);
                        return Ok(());
                    }
                }

                // class::member → look up class, then get static member
                self.compile_expr(class)?;
                if let ExprKind::Ident(name) = &member.kind {
                    let class_slot = self.define_local("__static_access_read_class");
                    self.emit_u16(Op::LOCAL_SET, class_slot);
                    let field_name = match &class.kind {
                        ExprKind::Ident(class_name) => {
                            self.js_member_storage_name_for_class(class_name, name)
                        }
                        _ => self.canon(name) };
                    if self.profile.supports_private_fields && name.starts_with('#') {
                        let getter_name = format!("__get_{}", field_name);
                        self.emit_u16(Op::LOCAL_GET, class_slot);
                        self.emit_const(Value::String(Arc::from(getter_name.as_str())));
                        // `has` (proto-walk, raw key) not `hasOwn`: the private
                        // accessor key is `__get_/__set___js_private_*` — a `__`
                        // key that `hasOwn` hides, and under prototype dispatch the
                        // accessor lives on the class prototype, not the instance.
                        let has_own_idx = self.import("ecma:object", "has");
                        let line = self.line;
                        self.emit_host_call(has_own_idx, 2);
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if_value(line);

                        self.emit_u16(Op::LOCAL_GET, class_slot);
                        let getter_key = self.str_const(&getter_name);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, getter_key);
                        self.emit_u16(Op::LOCAL_GET, class_slot);
                        self.emit_u8(Op::CALL_REF, 1);

                        self.chunk().emit_else(line);
                        self.emit_js_private_brand_check(class_slot, &field_name)?;
                        self.emit_u16(Op::LOCAL_GET, class_slot);
                        let idx = self.str_const(&field_name);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                        self.chunk().emit_end(line);
                    } else {
                        self.emit_u16(Op::LOCAL_GET, class_slot);
                        let idx = self.str_const(&field_name);
                        self.emit_struct_field_op(Op::STRUCT_GET, 0, idx);
                    }
                } else {
                    self.compile_expr(member)?;
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                }
            }

            // ── Match expression (PHP/Rust) ─────────────────────────────
            ExprKind::Match { subject, arms } => {
                self.compile_expr(subject)?;
                let subject_slot = self.define_local("__match_subj");
                self.emit_u16(Op::LOCAL_SET, subject_slot);
                let result_slot = self.define_local("__match_result");
                let matched_slot = self.define_local("__match_matched");
                self.emit_null();
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, matched_slot);

                for arm in arms {
                    self.emit_u16(Op::LOCAL_GET, matched_slot);
                    self.emit(Op::I32_EQZ);
                    let arm_line = self.line;
                    self.chunk().emit_if(arm_line);
                    if let Some(ref conditions) = arm.conditions {
                        let arm_match_slot = self.define_local("__match_arm_matches");
                        self.emit_const(Value::I32(0));
                        self.emit_u16(Op::LOCAL_SET, arm_match_slot);
                        for c in conditions {
                            self.emit_u16(Op::LOCAL_GET, subject_slot);
                            self.compile_expr(c)?;
                            {
                                let line = self.line;
                                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                            };
                            let cond_line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), cond_line);
                            self.chunk().emit_if(cond_line);
                            self.emit_const(Value::I32(1));
                            self.emit_u16(Op::LOCAL_SET, arm_match_slot);
                            self.chunk().emit_end(cond_line);
                        }
                        self.emit_u16(Op::LOCAL_GET, arm_match_slot);
                        let body_line = self.line;
                        self.chunk().emit_if(body_line);
                        self.compile_expr(&arm.body)?;
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                        self.chunk().emit_end(body_line);
                    } else {
                        // Default arm
                        self.compile_expr(&arm.body)?;
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                    }
                    self.chunk().emit_end(arm_line);
                }
                self.emit_u16(Op::LOCAL_GET, result_slot);
            }
        }
        Ok(())
    }

    /// `+`/`*`/`-` on two sets → union/intersection/difference.
    fn try_compile_set_arithmetic(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if !self.profile.set_arithmetic_operators {
            return Ok(false);
        }

        if self.expr_is_builtin_set(left) && self.expr_is_builtin_set(right) {
            let helper = match op {
                BinOp::Add => Some("__vybe_pascal_set_union"),
                BinOp::Mul => Some("__vybe_pascal_set_intersection"),
                BinOp::Sub => Some("__vybe_pascal_set_difference"),
                _ => None };

            if let Some(helper) = helper {
                self.emit_global_read(helper);
                self.compile_expr(left)?;
                self.compile_expr(right)?;
                self.emit_u8(Op::CALL_REF, 2);
                return Ok(true);
            }
        }

        let method_name = match op {
            BinOp::Add => "Add",
            BinOp::Eq | BinOp::NotEq => "Equal",
            _ => return Ok(false) };

        let Some(type_name) = self.pascal_binary_operator_type(left, right, method_name) else {
            return Ok(false);
        };

        let callee = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&type_name)),
            field: method_name.to_string(),
            null_safe: false });
        let args = vec![
            Argument::positional(left.clone()),
            Argument::positional(right.clone()),
        ];
        self.compile_call(&callee, &args)?;
        if *op == BinOp::NotEq {
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_not(self.chunk(), line);
            };
        }
        Ok(true)
    }

    fn try_compile_python_set_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if !self.profile.set_bitwise_operators {
            return Ok(false);
        }

        let helper = match op {
            BinOp::BitOr => Some("union"),
            BinOp::BitAnd => Some("intersection"),
            BinOp::Sub => Some("difference"),
            BinOp::BitXor => Some("symmetricDifference"),
            _ => None };

        let Some(helper) = helper else {
            return Ok(false);
        };

        self.compile_expr(left)?;
        self.compile_expr(right)?;
        let rhs_slot = self.define_local("__py_set_rhs");
        let lhs_slot = self.define_local("__py_set_lhs");
        self.emit_u16(Op::LOCAL_SET, rhs_slot);
        self.emit_u16(Op::LOCAL_SET, lhs_slot);

        let size_key = self.str_const("size");
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, size_key);
        self.emit(Op::REF_IS_NULL);
        let lhs_has_size_slot = self.define_local("__py_set_lhs_has_size");
        self.emit(Op::I32_EQZ);
        self.emit_u16(Op::LOCAL_SET, lhs_has_size_slot);

        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, size_key);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        self.emit_u16(Op::LOCAL_GET, lhs_has_size_slot);
        self.emit(Op::I32_AND);
        let line = self.line;
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        let idx = self.import("ecma:set", helper);
        self.emit_host_call(idx, 2);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.compile_binop(op);
        self.chunk().emit_end(line);
        Ok(true)
    }

    fn try_compile_csharp_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        // No language gate: the operands must share a declared type whose class
        // declares `op_Addition`/`op_Equality`/… — the CLR's mangled operator
        // spellings. A class only carries those names if a frontend emitted
        // them, so a language that does not speak them resolves nothing here
        // and the caller falls through to the ordinary binary path.
        let left_type = self.infer_expr_type_hint(left);
        let right_type = self.infer_expr_type_hint(right);
        let (Some(left_type), Some(right_type)) = (left_type, right_type) else {
            return Ok(false);
        };
        if Self::normalize_type_hint(&left_type) != Self::normalize_type_hint(&right_type) {
            return Ok(false);
        }

        let arg_exprs = [left, right];
        let (method_name, negate_result) = match op {
            BinOp::Add => ("op_Addition", false),
            BinOp::Eq => ("op_Equality", false),
            BinOp::NotEq
                if self.pending_class_has_method_name_for_type(&left_type, "op_Inequality") =>
            {
                ("op_Inequality", false)
            }
            BinOp::NotEq
                if self.pending_class_has_method_name_for_type(&left_type, "op_Equality") =>
            {
                ("op_Equality", true)
            }
            _ => return Ok(false) };

        if let Some(chunk_idx) =
            self.resolve_static_method_overload_chunk_for_type(&left_type, method_name, &arg_exprs)
        {
            self.emit_direct_static_method_call(chunk_idx, &arg_exprs)?;
            if negate_result {
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            }
            return Ok(true);
        }

        let Some(class_name) = self.resolve_pending_class_name_for_type_hint(&left_type) else {
            return Ok(false);
        };

        let callee = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&class_name)),
            field: method_name.to_string(),
            null_safe: false });
        let args = vec![
            Argument::positional(left.clone()),
            Argument::positional(right.clone()),
        ];
        self.compile_call(&callee, &args)?;
        if negate_result {
            {
                let line = self.line;
                crate::primitives::ops::emit_dyn_not(self.chunk(), line);
            };
        }
        Ok(true)
    }

    fn try_compile_fortran_interface_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        // Redundant name check removed: `resolve_fortran_operator_target`
        // resolves through the interface-overload map, empty for every other
        // profile.
        let arg_exprs = vec![left.clone(), right.clone()];
        let Some(target_name) = self.resolve_fortran_operator_target(op, &arg_exprs) else {
            return Ok(false);
        };

        let callee =
            if let Some(module_name) = self.enum_members.get(&self.canon(&target_name)).cloned() {
                let target_canon = self.canon(&target_name);
                let prefers_direct_module_global = self
                    .pending_classes
                    .get(&module_name)
                    .is_some_and(|pending| {
                        pending
                            .static_method_names
                            .iter()
                            .any(|member| member == &target_canon)
                    });
                if prefers_direct_module_global {
                    Expression::ident(&target_name)
                } else {
                    Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(&module_name)),
                        field: target_name.clone(),
                        null_safe: false })
                }
            } else {
                Expression::ident(&target_name)
            };
        let args = vec![
            Argument::positional(left.clone()),
            Argument::positional(right.clone()),
        ];
        self.compile_call(&callee, &args)?;
        Ok(true)
    }

    fn pascal_binary_operator_type(
        &self,
        left: &Expression,
        right: &Expression,
        method_name: &str,
    ) -> Option<String> {
        let left_type = self.pascal_expr_static_type(left)?;
        let right_type = self.pascal_expr_static_type(right)?;
        if !left_type.eq_ignore_ascii_case(&right_type) {
            return None;
        }

        let bare_type = left_type
            .split('<')
            .next()
            .unwrap_or(left_type.as_str())
            .trim();
        let canon_type = self.canon(bare_type);
        if !self.defined_globals.contains(&canon_type) {
            return None;
        }
        if !self
            .defined_class_methods
            .contains(&self.canon(method_name))
        {
            return None;
        }
        Some(bare_type.to_string())
    }

    pub(super) fn pascal_expr_static_type(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
            _ => None }
    }

    fn dotnet_expr_static_type(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
            ExprKind::New { class, .. } => match &class.kind {
                ExprKind::Ident(name) => Some(name.rsplit('.').next().unwrap_or(name).to_string()),
                ExprKind::Member { field, .. } => Some(field.to_string()),
                _ => None },
            ExprKind::Call { callee, .. } => match &callee.kind {
                ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("Parse") => {
                    match &object.kind {
                        ExprKind::Ident(name) if name.eq_ignore_ascii_case("Version") => {
                            Some("Version".into())
                        }
                        ExprKind::Member { field, .. } if field.eq_ignore_ascii_case("Version") => {
                            Some("Version".into())
                        }
                        _ => None }
                }
                _ => None },
            _ => None }
    }

    fn is_dotnet_type_name(type_name: &str, expected: &str) -> bool {
        type_name.eq_ignore_ascii_case(expected)
            || type_name
                .rsplit('.')
                .next()
                .is_some_and(|short| short.eq_ignore_ascii_case(expected))
    }

    fn try_compile_dotnet_datetime_timespan_binary_operator(
        &mut self,
        op: &BinOp,
        left: &Expression,
        right: &Expression,
    ) -> Result<bool, String> {
        if !self.profile.namespaces.use_dotnet {
            return Ok(false);
        }

        let Some(left_type) = self.dotnet_expr_static_type(left) else {
            return Ok(false);
        };
        let Some(right_type) = self.dotnet_expr_static_type(right) else {
            return Ok(false);
        };

        let emit = match op {
            BinOp::Add
                if Self::is_dotnet_type_name(&left_type, "TimeSpan")
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") =>
            {
                Some("dotnet.timespan_add")
            }
            BinOp::Sub
                if Self::is_dotnet_type_name(&left_type, "TimeSpan")
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") =>
            {
                Some("dotnet.timespan_sub")
            }
            BinOp::Add
                if Self::is_dotnet_type_name(&left_type, "DateTime")
                    && Self::is_dotnet_type_name(&right_type, "TimeSpan") =>
            {
                Some("dotnet.datetime_add_timespan")
            }
            BinOp::Sub
                if Self::is_dotnet_type_name(&left_type, "DateTime")
                    && Self::is_dotnet_type_name(&right_type, "DateTime") =>
            {
                Some("dotnet.datetime_subtract_datetime")
            }
            BinOp::Lt
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") =>
            {
                Some("dotnet.version_lt")
            }
            BinOp::Gt
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") =>
            {
                Some("dotnet.version_gt")
            }
            BinOp::Eq
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") =>
            {
                Some("dotnet.version_eq")
            }
            BinOp::NotEq
                if Self::is_dotnet_type_name(&left_type, "Version")
                    && Self::is_dotnet_type_name(&right_type, "Version") =>
            {
                Some("dotnet.version_ne")
            }
            _ => None };

        let Some(emit) = emit else {
            return Ok(false);
        };

        self.compile_expr(left)?;
        self.compile_expr(right)?;
        let line = self.line;
        self.emit_common(emit, 2, line);
        Ok(true)
    }

    pub(super) fn pascal_helper_function_name(&self, type_name: &str, method_name: &str) -> String {
        let sanitize = |text: &str| {
            text.chars()
                .map(|ch| {
                    if ch.is_ascii_alphanumeric() {
                        ch.to_ascii_lowercase()
                    } else {
                        '_'
                    }
                })
                .collect::<String>()
        };
        format!(
            "__pascal_helper_{}_{}",
            sanitize(type_name),
            sanitize(method_name),
        )
    }

    fn emit_c_unsigned_wrap(&mut self, modulus: f64) {
        self.emit_const(Value::F64(modulus));
        self.compile_binop(&BinOp::Mod);
        self.emit_const(Value::F64(modulus));
        self.emit(Op::F64_ADD);
        self.emit_const(Value::F64(modulus));
        self.compile_binop(&BinOp::Mod);
    }

    fn emit_c_signed_wrap_from_unsigned(&mut self, threshold: f64, modulus: f64) {
        inst!(self, core_wasm::dup);
        self.emit_const(Value::F64(threshold));
        self.emit(Op::F64_GE);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_const(Value::F64(modulus));
        self.emit(Op::F64_SUB);
        self.chunk().emit_else(line);
        self.chunk().emit_end(line);
    }
}

// ── Chunk-level expressions emit ────────────────────────────────────────────
// Free functions over `&mut Chunk`, merged in from the former `emitter::expressions`
// module. The `impl Compiler` walkers above and these primitives are the two
// halves of the same topic and now live in one file.
use crate::primitives::instructions::core_wasm;
use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
// Expression compilation helpers — shared bytecode patterns for common expressions.
//
// Ternary conditionals, short-circuit logic, and null coalescing are identical
// across all languages. Helpers emit structured WASM control constructs so
// callers can compile language-specific sub-expressions in between.

// ── Undefined sentinel ─────────────────────────────────────────────────
//
// JS `undefined` is not a WASM concept. We represent it as a global sentinel
// `__undefined` that is set up at bundle time. All compilers that need
// undefined semantics emit `global_get "__undefined"` via this helper.
// Languages that don't have undefined (VB, C#, Pascal, Python) never call this.
//
// This centralizes all undefined emission so the opcode `Op::undefined` can
// eventually be removed — every site that used to emit that opcode should
// call this function instead.

/// Emit the JS `undefined` value onto the stack.
/// Uses `global_get "__undefined"` — a sentinel wired at bundle time.
/// Stack: [] → [undefined]
pub fn emit_undefined(chunk: &mut Chunk, line: u32) {
    crate::primitives::globals::emit_read(chunk, "__undefined", line);
}

/// Emit bitwise NOT (i32). WASM equivalent: i32.const -1, i32.xor.
/// Stack: [i32] → [i32]
pub fn emit_i32_not(chunk: &mut Chunk, line: u32) {
    chunk.emit_i32_const(-1, line);
    chunk.emit_op(Op::I32_XOR, line);
}

/// Emit f64 C-style modulo as pure WASM opcodes (no host import).
/// Stack: [a, b] → [result]
pub fn emit_f64_mod_with_import(chunk: &mut Chunk, _import_idx: u16, line: u32) {
    crate::primitives::math::emit_c_fmod(chunk, line);
}

/// Emit f64 C-style modulo as pure WASM opcodes (no host import).
/// Stack: [a, b] → [result]
pub fn emit_f64_mod(chunk: &mut Chunk, line: u32) {
    crate::primitives::math::emit_c_fmod(chunk, line);
}

/// Emit boolean NOT. Converts value to bool then negates.
/// WASM equivalent: dyn_to_bool + i32.eqz.
/// Stack: [value] → [bool]
pub fn emit_bool_not(chunk: &mut Chunk, line: u32) {
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

// ── Ternary / conditional expression ────────────────────────────────────
//
// Usage:
//   compile_condition(chunk);
//   let false_jump = emit_ternary_start(chunk);
//   compile_then_expr(chunk);
//   let end_jump = emit_ternary_middle(chunk, false_jump);
//   compile_else_expr(chunk);
//   emit_ternary_end(chunk, end_jump);

/// After condition is on stack: convert to bool and enter the then arm.
/// Stack before: [condition]  Stack after: []
pub fn emit_ternary_start(chunk: &mut Chunk, line: u32) -> usize {
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    0
}

/// After "then" expression: start the else arm.
/// Stack: [then_value]
pub fn emit_ternary_middle(chunk: &mut Chunk, _false_jump: usize, line: u32) -> usize {
    chunk.emit_else(line);
    0
}

/// After "else" expression: close the structured if.
/// Stack: [result_value]
pub fn emit_ternary_end(chunk: &mut Chunk, _end_jump: usize) {
    chunk.emit_end(0);
}

// ── Short-circuit logical AND ───────────────────────────────────────────
//
// Usage:
//   compile_left(chunk);
//   let jump = emit_and_start(chunk);
//   compile_right(chunk);
//   emit_short_circuit_end(chunk, jump);

/// After left operand: if falsy, short-circuit (keep left as result).
/// Stack before: [left]  Stack after: [] (right will be compiled next)
pub fn emit_and_start(chunk: &mut Chunk, line: u32) -> usize {
    let block = chunk.emit_block(line);
    chunk.emit_dup(line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line); // discard left, right becomes result
    block
}

// ── Short-circuit logical OR ────────────────────────────────────────────
//
// Usage:
//   compile_left(chunk);
//   let jump = emit_or_start(chunk);
//   compile_right(chunk);
//   emit_short_circuit_end(chunk, jump);

/// After left operand: if truthy, short-circuit (keep left as result).
/// Stack before: [left]  Stack after: [] (right will be compiled next)
pub fn emit_or_start(chunk: &mut Chunk, line: u32) -> usize {
    let block = chunk.emit_block(line);
    chunk.emit_dup(line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line); // discard left, right becomes result
    block
}

/// End a short-circuit AND or OR.
/// Stack: [result_value]
pub fn emit_short_circuit_end(chunk: &mut Chunk, block: usize) {
    chunk.emit_end(0);
    chunk.patch_block(block);
}

// ── Null coalescing ─────────────────────────────────────────────────────
//
// Usage:
//   compile_left(chunk);
//   let (null_jump, end_jump) = emit_null_coalesce_start(chunk);
//   compile_right(chunk);
//   emit_null_coalesce_end(chunk, end_jump);

/// After left operand: if null, drop it and fall through to right expression.
/// If non-null, skip over right expression.
/// Stack before: [left]  Stack after: [] (right will be compiled next)
/// Returns (block_patch, 0).
pub fn emit_null_coalesce_start(chunk: &mut Chunk, line: u32) -> (usize, usize) {
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    (block, 0)
}

/// After right expression: close the non-null skip block.
/// Stack: [result_value]
pub fn emit_null_coalesce_end(chunk: &mut Chunk, block: usize) {
    chunk.emit_end(0);
    chunk.patch_block(block);
}

// ── Null-safe member access (?.) ────────────────────────────────────────
//
// Usage:
//   compile_object(chunk);
//   let (skip, end) = emit_null_safe_start(chunk);
//   // compile member access (struct_get, etc.)
//   emit_null_safe_end(chunk, end);

/// After object is on stack: if null, skip the member access.
/// Stack before: [object]  Stack after: [object] (if non-null) or control jumps to end
pub fn emit_null_safe_start(chunk: &mut Chunk, line: u32) -> (usize, usize) {
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let block = chunk.emit_block(line);
    chunk.emit_br_if(0, line);
    (block, 0)
}

/// After member access: close the null-skip block.
pub fn emit_null_safe_end(chunk: &mut Chunk, block: usize, _line: u32) {
    chunk.emit_end(0);
    chunk.patch_block(block);
}

// ── Generic dynamic dispatch ────────────────────────────────────────────
//
// The universal pattern for dynamic languages:
//   1. struct_get a method/property on the object
//   2. If found (non-null), call it
//   3. If not found, execute a fallback
//
// All rich_compare, smart_length, rich_arithmetic, etc. are instances of this.

/// Emit a try-method-or-fallback dispatch.
/// Checks if `obj_slot` has a method named `method_name`. If found, calls it with
/// `arg_count` args (which the caller pushes between start and end).
/// Returns (is_null_jump, found_done_jump) for the caller to emit the fallback.
///
/// Usage:
///   let (null_jump, done_jump) = emit_dynamic_dispatch_start(chunk, obj_slot, "method", line);
///   // push args for the found case
///   emit_dynamic_dispatch_call(chunk, arg_count, line);
///   let done = emit_dynamic_dispatch_middle(chunk, line);
///   // patch null case, emit fallback
///   emit_dynamic_dispatch_fallback(chunk, null_jump, done, line);
///
/// Or use the simpler one-shot helpers below.
// ── Rich arithmetic (user-defined __add__/__sub__/etc) ──────────────────
//
// Same pattern as rich_compare but for binary arithmetic operators.
// Tries user-defined dunder method, falls back to primitive opcode.

/// Emit rich arithmetic: tries the user-defined `__add__`/`__mul__`/…,
/// falls back to the primitive opcode.
/// Caller must store left in `left_slot` and right in `right_slot`.
/// Stack before: []  Stack after: [result_value]
///
/// Deliberately NOT `emit_rich_compare_locals`. That helper's second stage
/// tries `compare`/`CompareTo`/`<=>` and applies `fallback_fn` against `0`,
/// which is right for an ordering (`CompareTo(a, b) < 0`) and nonsense for
/// arithmetic (`CompareTo(a, b) + 0`). An operator method is the only thing
/// that can define `a + b`, so there is no second stage here.
pub fn emit_rich_arithmetic(
    chunk: &mut Chunk,
    left_slot: u16,
    right_slot: u16,
    dunder: &str,
    fallback_fn: fn(&mut Chunk, u32),
    line: u32,
) {
    let method_slot = alloc_local(chunk);

    // A primitive traps on STRUCT_GET, so gate the lookup on the operand
    // actually being an object.
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);

    let key = chunk.add_constant(Value::String(Arc::from(dunder)));
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    fallback_fn(chunk, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    fallback_fn(chunk, line);
    chunk.emit_end(line);
}

/// Emit a rich unary operator: tries the user-defined `__neg__`/`__bitnot__`
/// on an object operand, else the primitive fallback.
/// Stack before: []  Stack after: [result_value]
pub fn emit_rich_unary(
    chunk: &mut Chunk,
    operand_slot: u16,
    dunder: &str,
    fallback_fn: fn(&mut Chunk, u32),
    line: u32,
) {
    let method_slot = alloc_local(chunk);

    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    chunk.emit_op_u16(Op::LOCAL_GET, operand_slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);

    let key = chunk.add_constant(Value::String(Arc::from(dunder)));
    chunk.emit_op_u16(Op::LOCAL_GET, operand_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, operand_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, operand_slot, line);
    fallback_fn(chunk, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, operand_slot, line);
    fallback_fn(chunk, line);
    chunk.emit_end(line);
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }
    slot
}

// ── Rich toString (user-defined __str__ / toString) ─────────────────────

/// Emit smart toString: tries __str__/toString getter on objects, falls back to host toString.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [string_value]
pub fn emit_rich_to_string(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let method_slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(method_slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    // The ToString SLOT — filled by whatever the object's own language spells
    // it (`__str__`, `to_s`, `toString`, `__toString`), so this reaches a
    // user's string conversion regardless of where the class came from.
    let key = chunk.add_constant(Value::String(Arc::from(
        vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString).as_str(),
    )));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_call(to_str, 1, line);
    chunk.emit_end(line);
}

// ── Rich bool (user-defined __bool__ / valueOf) ─────────────────────────

/// Emit smart bool: tries __bool__ on objects, falls back to dyn_to_bool.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [bool_value]
pub fn emit_rich_bool(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let method_slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(method_slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from("__bool__")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_end(line);
}

// ── Rich comparison (user-defined __lt__/__gt__/etc) ────────────────────
//
// Standard WASM opcodes only. Emits inline dispatch:
//   1. Try struct_get for the dunder method on left operand
//   2. If found (non-null), call it with right operand
//   3. If not found, fall back to the primitive dyn_lt/dyn_gt/etc opcode
//
// This allows Python `__lt__`, Dart `operator<`, C# `CompareTo` etc.
// to work on user objects while keeping primitive comparison fast.

/// Emit a rich comparison: tries user-defined method, falls back to primitive opcode.
/// Both operands must already be on the stack: [left, right].
/// Stack after: [bool_result]
///
/// `dunder`: the method name to look for (e.g. "__lt__", "__gt__")
/// `fallback_fn`: the emitter to use if no method (e.g. `crate::primitives::ops::emit_dyn_lt`)
pub fn emit_rich_compare(
    chunk: &mut Chunk,
    _dunder: &str,
    fallback_fn: fn(&mut Chunk, u32),
    line: u32,
) {
    // Stack: [left, right]
    // Save right to temp, check left for dunder method
    // We need to: peek at left (under right), struct_get dunder, check null

    // Store right in temp
    // Note: we can't allocate locals here (no scope access). Use stack manipulation.
    // Strategy: swap to get left on top, dup, struct_get, check null.
    // But there's no swap opcode. Use a different approach:
    // Store right, dup left, struct_get dunder, check null.

    // Actually, the simplest approach that uses only standard WASM ops:
    // [left, right] on stack.
    // We need left for struct_get. But right is on top.
    // Emit: store right in a constant-indexed temp via the "over" pattern.

    // Simplest correct approach using only existing opcodes:
    // The caller must have left and right in locals already (common for binary ops).
    // But we take them from stack. Let's use the dup-under pattern:

    // For now, just use the fallback op. Rich compare requires local slots
    // which the caller must provide. Use emit_rich_compare_with_locals instead.
    fallback_fn(chunk, line);
}

/// Emit a rich comparison with pre-allocated local slots.
/// Caller must store left in `left_slot` and right in `right_slot` before calling.
/// Stack before: []  Stack after: [bool_result]
///
/// Emits: check left.__lt__ → if found, call it(right) → else dyn_lt(left, right)
/// What `emit_rich_compare_locals` falls back to when no user method matched.
///
/// `Op` is a plain chunk-only emitter — every relational operator uses one.
/// `Target` is an emit target from a `[builtin_slots.*] eq` binding, which is
/// how a language declares STRUCTURAL equality for its composite built-ins
/// (Python's order-independent set equality, Dart's record/tuple equality).
/// Those two used to arrive as `LanguageHooks::value_eq`, looked up by language
/// NAME in shared code — builtinslotplan.md §3c.
///
/// A target needs the whole chunk vector to dispatch, which is the only reason
/// this function takes `chunks`/`current` rather than a single chunk.
pub enum RichFallback<'a> {
    Op(fn(&mut Chunk, u32)),
    Target(&'a str) }

impl RichFallback<'_> {
    fn emit(&self, chunks: &mut Vec<Chunk>, current: usize, line: u32) {
        match self {
            RichFallback::Op(f) => f(&mut chunks[current], line),
            RichFallback::Target(target) => {
                if let Some(name) = target.strip_prefix("common:") {
                    // A miss emits NOTHING, stranding both operands on the
                    // stack for whatever consumes the result next — the same
                    // silent corruption the `compare` targets guard against.
                    assert!(
                        crate::primitives::dispatch::emit_common(name, chunks, current, 2, line),
                        "[builtin_slots] eq target `common:{name}` is not dispatched"
                    );
                } else if let Some(rest) = target.strip_prefix("host:") {
                    let (module, func) = rest.rsplit_once(':').unwrap_or_else(|| {
                        panic!("[builtin_slots] eq target `{target}` is not `host:<m>:<fn>`")
                    });
                    let idx = chunks[current].add_import(module, func);
                    chunks[current].emit_call(idx, 2, line);
                } else {
                    panic!("[builtin_slots] eq target `{target}` must be `common:…` or `host:…`");
                }
            }
        }
    }
}

pub fn emit_rich_compare_locals(
    chunks: &mut Vec<Chunk>,
    current: usize,
    left_slot: u16,
    right_slot: u16,
    dunder: &str,
    fallback: RichFallback<'_>,
    line: u32,
) {
    // Re-ASSIGNED (not shadowed) around each `fallback.emit`, which needs the
    // whole vector: shadowing would leave the outer borrow live across the
    // loop's next iteration.
    let mut chunk = &mut chunks[current];
    let method_slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(method_slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }

    // Try struct_get dunder on left
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from(dunder)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    // Found method: call it with self=left, arg=right → result
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_else(line);

    // Not found: try compare-style methods like C# CompareTo / Ruby <=>.
    // The block CARRIES the comparison result (one value on every path):
    // the `br 1` below targets it, and a branch to a void block would
    // discard the value it just computed — the outer consumer then read
    // whatever sat beneath it on the stack (measured: Null into an `if`).
    let done = chunk.emit_block_typed(line, 1);
    for method_name in ["compare", "CompareTo", "compareTo", "__cmp__", "<=>"] {
        let method_key = chunk.add_constant(Value::String(Arc::from(method_name)));
        chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, 0, method_key, line);
        chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

        chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
        chunk.emit_op_u8(Op::CALL_REF, 2, line);
        core_wasm::i32_const(chunk, line, 0);
        fallback.emit(chunks, current, line);
        chunk = &mut chunks[current];
        chunk.emit_br(1, line);
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    fallback.emit(chunks, current, line);
    chunk = &mut chunks[current];

    chunk.emit_end(line);
    chunk.patch_block(done);
    chunk.emit_end(line);
}

// ── Smart length (user-defined __len__ / __get_length) ──────────────────
//
// Standard WASM opcodes only. Tries __get_length getter first,
// falls back to array_length opcode for plain arrays/strings.

/// Emit smart length: tries user-defined __get_length getter, falls back to array_length.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [length_value]
pub fn emit_smart_length(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let method_slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(method_slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }

    // The Len SLOT first — Python `__len__`, Ruby `size`, C# `Count`, Java
    // `size()` all fill it, so a user-defined length works whatever language
    // declared the class.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let slot_key = chunk.add_constant(Value::String(Arc::from(
        vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Len).as_str(),
    )));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, slot_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    // Then the property form. This probe also serves values that never went
    // through a class — a `length` PROPERTY lowers to a `__get_length` getter
    // and carries no slot, and a plain array carries neither. So the slot is an
    // additional key here, not a replacement: substituting it would fix
    // `__len__` and break every property-based length.
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from("__get_length")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_end(line);
}
