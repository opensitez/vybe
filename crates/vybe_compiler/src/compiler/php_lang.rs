use vybe_ast::{ExprKind, Expression, Literal};
use crate::compiler::*;
use vybe_emitter as common;
use std::sync::Arc;
use vybe_bytecode::{Op, Value};

#[derive(Clone, Copy)]
enum BufferedGeneratorStepMode {
    Valid,
    Current,
    Value,
}

impl Compiler {
    fn emit_buffered_generator_set_bool_property(&mut self, obj_slot: u16, key: u16, value: bool) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::Bool(value));
        self.emit_u16(Op::STRUCT_SET, key);
        self.emit(Op::DROP);
    }

    fn emit_buffered_generator_mark_started(&mut self, obj_slot: u16, started_key: u16) {
        self.emit_buffered_generator_set_bool_property(obj_slot, started_key, true);
    }

    fn emit_buffered_generator_store_yielded_state(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        done_key: u16,
        current_key: u16,
    ) {
        self.emit_buffered_generator_set_bool_property(obj_slot, done_key, false);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_generator_yield_value(value_slot);
        self.emit_u16(Op::STRUCT_SET, current_key);
        self.emit(Op::DROP);
    }

    fn emit_buffered_generator_store_completed_state(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_buffered_generator_set_bool_property(obj_slot, done_key, true);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_u16(Op::STRUCT_SET, return_key);
        self.emit(Op::DROP);
        self.emit_buffered_generator_set_bool_property(obj_slot, current_key, false);
    }

    fn emit_buffered_generator_set_step_result(
        &mut self,
        value_slot: u16,
        result_slot: u16,
        mode: BufferedGeneratorStepMode,
        yielded: bool,
    ) {
        match (mode, yielded) {
            (BufferedGeneratorStepMode::Valid, true) => self.emit_const(Value::Bool(true)),
            (BufferedGeneratorStepMode::Valid, false) => self.emit_const(Value::Bool(false)),
            (BufferedGeneratorStepMode::Current, false) => self.emit_const(Value::Bool(false)),
            (BufferedGeneratorStepMode::Current, true)
            | (BufferedGeneratorStepMode::Value, true) => {
                self.emit_generator_yield_value(value_slot);
            }
            (BufferedGeneratorStepMode::Value, false) => {
                self.emit_u16(Op::LOCAL_GET, value_slot);
            }
        }
        self.emit_u16(Op::LOCAL_SET, result_slot);
    }

    fn emit_buffered_generator_apply_next_result(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        has_more_slot: u16,
        result_slot: u16,
        mode: BufferedGeneratorStepMode,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_u16(Op::LOCAL_GET, has_more_slot);
        {
            let line = self.line;
            vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_buffered_generator_store_yielded_state(
            obj_slot,
            value_slot,
            done_key,
            current_key,
        );
        self.emit_buffered_generator_set_step_result(value_slot, result_slot, mode, true);

        self.chunk().emit_else(line);

        self.emit_buffered_generator_store_completed_state(
            obj_slot,
            value_slot,
            done_key,
            current_key,
            return_key,
        );
        self.emit_buffered_generator_set_step_result(value_slot, result_slot, mode, false);

        self.chunk().emit_end(line);
    }

    fn emit_buffered_generator_apply_resume_result(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        result_slot: u16,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let is_done_idx = self.import("ecma:value", "isGeneratorDone");
        self.emit_host_call(is_done_idx, 1);
        {
            let line = self.line;
            vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_buffered_generator_store_completed_state(
            obj_slot,
            value_slot,
            done_key,
            current_key,
            return_key,
        );
        self.emit_buffered_generator_set_step_result(
            value_slot,
            result_slot,
            BufferedGeneratorStepMode::Value,
            false,
        );

        self.chunk().emit_else(line);

        self.emit_buffered_generator_store_yielded_state(
            obj_slot,
            value_slot,
            done_key,
            current_key,
        );
        self.emit_buffered_generator_set_step_result(
            value_slot,
            result_slot,
            BufferedGeneratorStepMode::Value,
            true,
        );

        self.chunk().emit_end(line);
    }

    fn emit_buffered_generator_start_with_next(
        &mut self,
        obj_slot: u16,
        result_slot: u16,
        mode: BufferedGeneratorStepMode,
        started_key: u16,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let line = self.line;
        vybe_emitter::generators::emit_next(self.chunk(), line);
        let has_more_slot = self.define_local("__php_gen_has_more");
        self.emit_u16(Op::LOCAL_SET, has_more_slot);
        let value_slot = self.define_local("__php_gen_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.emit_buffered_generator_mark_started(obj_slot, started_key);
        self.emit_buffered_generator_apply_next_result(
            obj_slot,
            value_slot,
            has_more_slot,
            result_slot,
            mode,
            done_key,
            current_key,
            return_key,
        );
    }

    pub(crate) fn maybe_define_buffered_generator_key_index_slot(
        &mut self,
        key: Option<&str>,
    ) -> Option<u16> {
        if self.profile.buffered_iterator_methods && key.is_some() {
            let slot = self.define_local("__php_gen_loop_index");
            self.emit_const(Value::F64(0.0));
            self.emit_u16(Op::LOCAL_SET, slot);
            Some(slot)
        } else {
            None
        }
    }

    pub(crate) fn emit_buffered_generator_foreach_state(
        &mut self,
        cont_slot: u16,
        has_more_slot: u16,
        value_slot: u16,
    ) {
        let started_key = self.str_const("__php_gen_started");
        let current_key = self.str_const("__php_gen_current");
        let done_key = self.str_const("__php_gen_done");
        let return_key = self.str_const("__php_gen_return");

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(true));
        self.emit_u16(Op::STRUCT_SET, started_key);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, has_more_slot);
        {
            let line = self.line;
            vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(false));
        self.emit_u16(Op::STRUCT_SET, done_key);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_generator_yield_value(value_slot);
        self.emit_u16(Op::STRUCT_SET, current_key);
        self.emit(Op::DROP);
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(true));
        self.emit_u16(Op::STRUCT_SET, done_key);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_u16(Op::STRUCT_SET, return_key);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(false));
        self.emit_u16(Op::STRUCT_SET, current_key);
        self.emit(Op::DROP);
        self.chunk().emit_br(2, line);

        self.chunk().emit_end(line);
    }

    pub(crate) fn emit_buffered_generator_key_binding(
        &mut self,
        key_slot: u16,
        value_slot: u16,
        key_index_slot: Option<u16>,
    ) {
        self.emit_generator_yield_key_or_fallback(value_slot, key_index_slot);
        self.emit_u16(Op::LOCAL_SET, key_slot);
    }

    pub(crate) fn emit_buffered_generator_value_binding(&mut self, var_slot: u16, value_slot: u16) {
        self.emit_generator_yield_value(value_slot);
        self.emit_u16(Op::LOCAL_SET, var_slot);
    }

    pub(crate) fn emit_buffered_generator_method_dispatch(
        &mut self,
        obj_tmp: u16,
        field_name: &str,
        arg_exprs: &[&Expression],
    ) -> Result<Option<usize>, String> {
        let is_buffered_generator_method = (field_name == "current" && arg_exprs.is_empty())
            || (field_name == "send" && arg_exprs.len() == 1)
            || (field_name == "next" && arg_exprs.is_empty())
            || (field_name == "throw" && arg_exprs.len() == 1)
            || (field_name == "valid" && arg_exprs.is_empty())
            || (field_name == "getReturn" && arg_exprs.is_empty())
            || (field_name == "rewind" && arg_exprs.is_empty());

        if !is_buffered_generator_method {
            return Ok(None);
        }

        let started_key = self.str_const("__php_gen_started");
        let current_key = self.str_const("__php_gen_current");
        let done_key = self.str_const("__php_gen_done");
        let return_key = self.str_const("__php_gen_return");
        let result_slot = self.define_local("__php_gen_method_result");

        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        let is_gen_idx = self.import("ecma:value", "isGenerator");
        self.emit_host_call(is_gen_idx, 1);
        let line = self.line;
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        match field_name {
            "getReturn" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, return_key);
                self.emit_u16(Op::LOCAL_SET, result_slot);
            }
            "valid" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_not(self.chunk(), line);
                };
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_buffered_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    BufferedGeneratorStepMode::Valid,
                    started_key,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);
            }
            "current" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, current_key);
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_end(line);
                self.chunk().emit_else(line);
                self.emit_buffered_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    BufferedGeneratorStepMode::Current,
                    started_key,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);
            }
            "send" | "next" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                if field_name == "send" {
                    self.compile_expr(&arg_exprs[0])?;
                } else {
                    self.emit(Op::NULL);
                }
                let line = self.line;
                vybe_emitter::generators::emit_resume(self.chunk(), line);
                let value_slot = self.define_local("__php_gen_resume_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit_buffered_generator_apply_resume_result(
                    obj_tmp,
                    value_slot,
                    result_slot,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);

                self.chunk().emit_else(line);
                self.emit_buffered_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    BufferedGeneratorStepMode::Value,
                    started_key,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);
                // Mark as moved for rewind() check
                let moved_key = self.str_const("__php_gen_moved");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, moved_key);
                self.emit(Op::DROP);
            }
            "throw" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(&arg_exprs[0])?;
                let line = self.line;
                vybe_emitter::generators::emit_resume_throw(self.chunk(), line);
                let value_slot = self.define_local("__php_gen_throw_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit_buffered_generator_apply_resume_result(
                    obj_tmp,
                    value_slot,
                    result_slot,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let line = self.line;
                vybe_emitter::generators::emit_next(self.chunk(), line);
                let has_more_slot = self.define_local("__php_gen_throw_has_more");
                self.emit_u16(Op::LOCAL_SET, has_more_slot);
                let start_value_slot = self.define_local("__php_gen_throw_start_value");
                self.emit_u16(Op::LOCAL_SET, start_value_slot);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, started_key);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(&arg_exprs[0])?;
                let line = self.line;
                vybe_emitter::generators::emit_resume_throw(self.chunk(), line);
                let start_resume_slot = self.define_local("__php_gen_throw_resume_value");
                self.emit_u16(Op::LOCAL_SET, start_resume_slot);
                self.emit_buffered_generator_apply_resume_result(
                    obj_tmp,
                    start_resume_slot,
                    result_slot,
                    done_key,
                    current_key,
                    return_key,
                );

                self.chunk().emit_else(line);
                self.emit_buffered_generator_store_completed_state(
                    obj_tmp,
                    start_value_slot,
                    done_key,
                    current_key,
                    return_key,
                );
                self.emit_buffered_generator_set_step_result(
                    start_value_slot,
                    result_slot,
                    BufferedGeneratorStepMode::Value,
                    false,
                );

                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
            }
            "rewind" => {
                // PHP Generator::rewind() throws if the generator has
                // been advanced past the initial yield (via next/send).
                let moved_key = self.str_const("__php_gen_moved");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, moved_key);
                {
                    let line = self.line;
                    vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_const(Value::String(Arc::from(
                    "Cannot rewind a generator that was already run",
                )));
                common::errors::emit_throw(&mut self.chunks[self.current], line);
                self.chunk().emit_end(line);
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, result_slot);
            }
            _ => unreachable!(),
        }

        let line = self.line;
        self.chunk().emit_else(line);
        Ok(Some(result_slot as usize))
    }

    pub(crate) fn emit_buffered_generator_close_ident_if_needed(&mut self, name: &str) {
        if !self.profile.buffered_iterator_methods {
            return;
        }

        self.emit_var_get(name);
        let gen_slot = self.define_local("__buffered_generator_overwrite");
        self.emit_u16(Op::LOCAL_SET, gen_slot);

        self.emit_u16(Op::LOCAL_GET, gen_slot);
        let is_generator = self.import("ecma:value", "isGenerator");
        self.emit_host_call(is_generator, 1);
        let line = self.line;
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, gen_slot);
        let started_key = self.str_const("__php_gen_started");
        self.emit_u16(Op::STRUCT_GET, started_key);
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, gen_slot);
        let done_key = self.str_const("__php_gen_done");
        self.emit_u16(Op::STRUCT_GET, done_key);
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, gen_slot);
        self.emit(Op::NULL);
        self.emit_generator_control_packet_from_stack("return");
        let line = self.line;
        vybe_emitter::generators::emit_resume(self.chunk(), line);
        self.emit(Op::DROP);

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
    }
    pub(crate) fn emit_php_dynamic_function_name_resolution(&mut self, callee_slot: u16) {
        if !self.is_php_profile() {
            return;
        }

        let mut known_functions: Vec<String> = self.defined_functions.iter().cloned().collect();
        if known_functions.is_empty() {
            return;
        }
        known_functions.sort();

        self.emit_u16(Op::LOCAL_GET, callee_slot);
        {
            let l = self.line;
            vybe_emitter::instructions::host::CapabilityContext::get()
                .functions
                .emit(&mut self.chunks[self.current], "ecma:value", "typeof", 1, l);
        };
        self.emit_const(Value::String(Arc::from("string")));
        {
            let line = self.line;
            vybe_emitter::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, callee_slot);
        let line = self.line;
        common::strings::emit_to_lower(self.chunk(), line);
        let callee_name_slot = self.define_local("__php_string_callee_name");
        self.emit_u16(Op::LOCAL_SET, callee_name_slot);

        let matched_slot = self.define_local("__php_string_callee_matched");
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        for function_name in known_functions {
            let lowered_name = function_name.to_ascii_lowercase();
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, callee_name_slot);
            self.emit_const(Value::String(Arc::from(lowered_name.as_str())));
            {
                let line = self.line;
                vybe_emitter::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);

            let idx = self.str_const(&function_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u16(Op::LOCAL_SET, callee_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }
        self.chunk().emit_end(line);
    }

    /// Resolve a dynamic class-name string (`new $c`, `$cls::method()`) to its
    /// constructor reference at runtime. Mirrors
    /// [`Self::emit_php_dynamic_function_name_resolution`]: emit a compile-time
    /// match chain over the known classes (class names are case-insensitive in
    /// PHP, so compare lowercased) and, on match, replace the value in
    /// `class_slot` with the class's constructor ref. A non-string value (an
    /// already-resolved class object, e.g. `new static`/`new $obj`) is left
    /// untouched. `ctor_arity` = `Some(n)` resolves to the arity-`n`-specialised
    /// *constructor* global (for `new $c`); `None` resolves to the *class
    /// object* global (for `$cls::staticMethod()` — static members live on it).
    pub(crate) fn emit_php_dynamic_class_name_resolution(
        &mut self,
        class_slot: u16,
        ctor_arity: Option<usize>,
    ) {
        if !self.is_php_profile() {
            return;
        }
        let mut known_classes: Vec<String> = self.defined_classes.iter().cloned().collect();
        if known_classes.is_empty() {
            return;
        }
        known_classes.sort();

        self.emit_u16(Op::LOCAL_GET, class_slot);
        {
            let l = self.line;
            vybe_emitter::instructions::host::CapabilityContext::get()
                .functions
                .emit(&mut self.chunks[self.current], "ecma:value", "typeof", 1, l);
        };
        self.emit_const(Value::String(Arc::from("string")));
        {
            let line = self.line;
            vybe_emitter::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, class_slot);
        let line = self.line;
        common::strings::emit_to_lower(self.chunk(), line);
        // Class names may carry a leading namespace separator; the ctor globals
        // are stored under the bare name, so strip it for matching too.
        let class_name_slot = self.define_local("__php_string_class_name");
        self.emit_u16(Op::LOCAL_SET, class_name_slot);

        let matched_slot = self.define_local("__php_string_class_matched");
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        for class_name in known_classes {
            let bare = Self::strip_global_namespace_prefix(&class_name).to_string();
            let lowered_name = bare.to_ascii_lowercase();
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, class_name_slot);
            self.emit_const(Value::String(Arc::from(lowered_name.as_str())));
            {
                let line = self.line;
                vybe_emitter::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);

            let canon = self.canon(&bare);
            match ctor_arity {
                Some(arity) => {
                    let primary_ctor = format!("{}$arity{}", canon, arity);
                    self.emit_dynamic_constructor_global_ref(&primary_ctor, Some(&canon), &bare);
                }
                None => {
                    let idx = self.str_const(&canon);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                }
            }
            self.emit_u16(Op::LOCAL_SET, class_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }
        self.chunk().emit_end(line);
    }

    pub(crate) fn finish_buffered_generator_method_dispatch(&mut self, result_slot: usize) {
        let line = self.line;
        self.emit_u16(Op::LOCAL_SET, result_slot as u16);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, result_slot as u16);
    }
    pub(crate) fn compile_php_autoload_callable_ref(
        &mut self,
        expr: &Expression,
    ) -> Result<(), String> {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(function_name)) => {
                let resolved_name = self.resolve_source_type_alias(function_name);
                let function_idx = self.str_const(&self.canon(&resolved_name));
                self.emit_u16(Op::GLOBAL_GET, function_idx);
                Ok(())
            }
            _ => self.compile_expr(expr),
        }
    }
    pub(crate) fn resolve_php_autoload_callback_class_global(
        &self,
        class_name: &str,
    ) -> Option<String> {
        let resolved_class = self.resolve_source_type_alias(class_name);
        let canon_class = self.canon(&resolved_class);
        if self.defined_classes.contains(&canon_class)
            || self.defined_globals.contains(&canon_class)
        {
            return Some(canon_class);
        }
        resolved_class.rsplit('.').next().and_then(|short_name| {
            let short_canon = self.canon(short_name);
            if self.defined_classes.contains(&short_canon)
                || self.defined_globals.contains(&short_canon)
            {
                Some(short_canon)
            } else {
                None
            }
        })
    }
    pub(crate) fn is_php_profile(&self) -> bool {
        self.profile.name == "php"
    }
    pub(crate) fn emit_php_promote_empty_array_for_string_key(
        &mut self,
        obj_slot: u16,
        key_slot: u16,
        line: u32,
    ) {
        let is_array_idx = self.import("ecma:array", "isArray");
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.chunk()
            .emit_op_u16(Op::CALL_IMPORT, is_array_idx, line);
        self.chunk().emit(1, line);
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, key_slot);
        {
            let l = self.line;
            vybe_emitter::instructions::host::CapabilityContext::get()
                .functions
                .emit(
                    &mut self.chunks[self.current],
                    "wasm:js-string",
                    "test",
                    1,
                    l,
                );
        };
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit(Op::ARRAY_LENGTH);
        {
            let l = self.line;
            vybe_emitter::instructions::core_wasm::i32_const(
                &mut self.chunks[self.current],
                l,
                0,
            );
        };
        {
            let line = self.line;
            vybe_emitter::ops::emit_dyn_ne(self.chunk(), line);
        };
        vybe_emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);
        let map_new_idx = self.import("ecma:map", "new");
        self.chunk().emit_op_u16(Op::CALL_IMPORT, map_new_idx, line);
        self.chunk().emit(0, line);
        self.emit_u16(Op::LOCAL_SET, obj_slot);

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
    }
}
