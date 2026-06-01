use super::super::compiler::*;
use super::super::ast::Expression;
use vybe_bytecode::Value;

#[derive(Clone, Copy)]
enum PhpGeneratorStepMode {
    Valid,
    Current,
    Value,
}

impl Compiler {
    fn emit_php_generator_set_bool_property(&mut self, obj_slot: u16, key: u16, value: bool) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::Bool(value));
        self.emit_u16(Op::STRUCT_SET, key);
        self.emit(Op::DROP);
    }

    fn emit_php_generator_mark_started(&mut self, obj_slot: u16, started_key: u16) {
        self.emit_php_generator_set_bool_property(obj_slot, started_key, true);
    }

    fn emit_php_generator_store_yielded_state(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        done_key: u16,
        current_key: u16,
    ) {
        self.emit_php_generator_set_bool_property(obj_slot, done_key, false);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_generator_yield_value(value_slot);
        self.emit_u16(Op::STRUCT_SET, current_key);
        self.emit(Op::DROP);
    }

    fn emit_php_generator_store_completed_state(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_php_generator_set_bool_property(obj_slot, done_key, true);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_u16(Op::STRUCT_SET, return_key);
        self.emit(Op::DROP);
        self.emit_php_generator_set_bool_property(obj_slot, current_key, false);
    }

    fn emit_php_generator_set_step_result(
        &mut self,
        value_slot: u16,
        result_slot: u16,
        mode: PhpGeneratorStepMode,
        yielded: bool,
    ) {
        match (mode, yielded) {
            (PhpGeneratorStepMode::Valid, true) => self.emit_const(Value::Bool(true)),
            (PhpGeneratorStepMode::Valid, false) => self.emit_const(Value::Bool(false)),
            (PhpGeneratorStepMode::Current, false) => self.emit_const(Value::Bool(false)),
            (PhpGeneratorStepMode::Current, true) | (PhpGeneratorStepMode::Value, true) => {
                self.emit_generator_yield_value(value_slot);
            }
            (PhpGeneratorStepMode::Value, false) => {
                self.emit_u16(Op::LOCAL_GET, value_slot);
            }
        }
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit(Op::DROP);
    }

    fn emit_php_generator_apply_next_result(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        has_more_slot: u16,
        result_slot: u16,
        mode: PhpGeneratorStepMode,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_u16(Op::LOCAL_GET, has_more_slot);
        { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_php_generator_store_yielded_state(obj_slot, value_slot, done_key, current_key);
        self.emit_php_generator_set_step_result(value_slot, result_slot, mode, true);

        self.chunk().emit_else(line);

        self.emit_php_generator_store_completed_state(obj_slot, value_slot, done_key, current_key, return_key);
        self.emit_php_generator_set_step_result(value_slot, result_slot, mode, false);

        self.chunk().emit_end(line);
    }

    fn emit_php_generator_apply_resume_result(
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
        { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_php_generator_store_completed_state(obj_slot, value_slot, done_key, current_key, return_key);
        self.emit_php_generator_set_step_result(value_slot, result_slot, PhpGeneratorStepMode::Value, false);

        self.chunk().emit_else(line);

        self.emit_php_generator_store_yielded_state(obj_slot, value_slot, done_key, current_key);
        self.emit_php_generator_set_step_result(value_slot, result_slot, PhpGeneratorStepMode::Value, true);

        self.chunk().emit_end(line);
    }

    fn emit_php_generator_start_with_next(
        &mut self,
        obj_slot: u16,
        result_slot: u16,
        mode: PhpGeneratorStepMode,
        started_key: u16,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit(Op::GEN_NEXT);
        let has_more_slot = self.define_local("__php_gen_has_more");
        self.emit_u16(Op::LOCAL_SET, has_more_slot);
        self.emit(Op::DROP);
        let value_slot = self.define_local("__php_gen_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);

        self.emit_php_generator_mark_started(obj_slot, started_key);
        self.emit_php_generator_apply_next_result(
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

    pub(super) fn maybe_define_php_generator_key_index_slot(&mut self, key: Option<&str>) -> Option<u16> {
        if self.is_php_profile() && key.is_some() {
            let slot = self.define_local("__php_gen_loop_index");
            self.emit_const(Value::F64(0.0));
            self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
            Some(slot)
        } else {
            None
        }
    }

    pub(super) fn emit_php_generator_foreach_state(&mut self, cont_slot: u16, has_more_slot: u16, value_slot: u16) {
        let started_key = self.str_const("__php_gen_started");
        let current_key = self.str_const("__php_gen_current");
        let done_key = self.str_const("__php_gen_done");
        let return_key = self.str_const("__php_gen_return");

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(true));
        self.emit_u16(Op::STRUCT_SET, started_key);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, has_more_slot);
        { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
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

    pub(super) fn emit_php_generator_key_binding(&mut self, key_slot: u16, value_slot: u16, key_index_slot: Option<u16>) {
        self.emit_generator_yield_key_or_fallback(value_slot, key_index_slot);
        self.emit_u16(Op::LOCAL_SET, key_slot); self.emit(Op::DROP);
    }

    pub(super) fn emit_php_generator_value_binding(&mut self, var_slot: u16, value_slot: u16) {
        self.emit_generator_yield_value(value_slot);
        self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);
    }

    pub(super) fn emit_php_generator_method_dispatch(
        &mut self,
        obj_tmp: u16,
        field_name: &str,
        arg_exprs: &[&Expression],
    ) -> Result<Option<usize>, String> {
        let is_php_generator_method = (field_name == "current" && arg_exprs.is_empty())
            || (field_name == "send" && arg_exprs.len() == 1)
            || (field_name == "next" && arg_exprs.is_empty())
            || (field_name == "throw" && arg_exprs.len() == 1)
            || (field_name == "valid" && arg_exprs.is_empty())
            || (field_name == "getReturn" && arg_exprs.is_empty());

        if !is_php_generator_method {
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
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        match field_name {
            "getReturn" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, return_key);
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);
            }
            "valid" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                { let line = self.line; crate::emitter::ops::emit_dyn_not(self.chunk(), line); };
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);

                self.chunk().emit_else(line);
                self.emit_php_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    PhpGeneratorStepMode::Valid,
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
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, current_key);
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);

                self.chunk().emit_end(line);
                self.chunk().emit_else(line);
                self.emit_php_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    PhpGeneratorStepMode::Current,
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
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                if field_name == "send" {
                    self.compile_expr(&arg_exprs[0])?;
                } else {
                    self.emit(Op::NULL);
                }
                self.emit_u16(Op::RESUME, 0);
                let value_slot = self.define_local("__php_gen_resume_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                self.emit_php_generator_apply_resume_result(
                    obj_tmp,
                    value_slot,
                    result_slot,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);

                self.chunk().emit_else(line);
                self.emit_php_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    PhpGeneratorStepMode::Value,
                    started_key,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);
            }
            "throw" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(&arg_exprs[0])?;
                self.emit_generator_control_packet_from_stack("throw");
                self.emit_u16(Op::RESUME, 0);
                let value_slot = self.define_local("__php_gen_throw_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                self.emit_php_generator_apply_resume_result(
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
                self.emit(Op::GEN_NEXT);
                let has_more_slot = self.define_local("__php_gen_throw_has_more");
                self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                let start_value_slot = self.define_local("__php_gen_throw_start_value");
                self.emit_u16(Op::LOCAL_SET, start_value_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, started_key);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(&arg_exprs[0])?;
                self.emit_generator_control_packet_from_stack("throw");
                self.emit_u16(Op::RESUME, 0);
                let start_resume_slot = self.define_local("__php_gen_throw_resume_value");
                self.emit_u16(Op::LOCAL_SET, start_resume_slot); self.emit(Op::DROP);
                self.emit_php_generator_apply_resume_result(
                    obj_tmp,
                    start_resume_slot,
                    result_slot,
                    done_key,
                    current_key,
                    return_key,
                );

                self.chunk().emit_else(line);
                self.emit_php_generator_store_completed_state(
                    obj_tmp,
                    start_value_slot,
                    done_key,
                    current_key,
                    return_key,
                );
                self.emit_php_generator_set_step_result(
                    start_value_slot,
                    result_slot,
                    PhpGeneratorStepMode::Value,
                    false,
                );

                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
            }
            _ => unreachable!(),
        }

        let line = self.line;
        self.chunk().emit_else(line);
        Ok(Some(result_slot as usize))
    }
}
