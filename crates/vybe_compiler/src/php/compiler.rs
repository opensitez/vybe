use super::super::compiler::*;
use super::super::ast::Expression;
use vybe_bytecode::Value;
use vybe_bytecode::opcode::Op;

impl Compiler {
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
        let exhausted = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(false));
        self.emit_u16(Op::STRUCT_SET, done_key);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_generator_yield_value(value_slot);
        self.emit_u16(Op::STRUCT_SET, current_key);
        self.emit(Op::DROP);
        let loop_ready = self.emit_jump(Op::BR);

        self.patch_jump(exhausted);
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
        self.emit_u8(Op::BR_LABEL, 1);

        self.patch_jump(loop_ready);
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

        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        let is_gen_idx = self.import("ecma:value", "isGenerator");
        self.emit_host_call(is_gen_idx, 1);
        let not_gen = self.emit_jump(Op::BR_IF_FALSE);

        match field_name {
            "getReturn" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, return_key);
            }
            "valid" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let need_start = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                { let line = self.line; crate::emitter::ops::emit_dyn_not(self.chunk(), line); };
                let handled = self.emit_jump(Op::BR);

                self.patch_jump(need_start);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit(Op::GEN_NEXT);
                let has_more_slot = self.define_local("__php_gen_has_more");
                self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                let value_slot = self.define_local("__php_gen_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, started_key);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let no_more = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_generator_yield_value(value_slot);
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_const(Value::Bool(true));
                let start_done = self.emit_jump(Op::BR);

                self.patch_jump(no_more);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_u16(Op::STRUCT_SET, return_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_const(Value::Bool(false));

                self.patch_jump(handled);
                self.patch_jump(start_done);
            }
            "current" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let need_start = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let not_done = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_const(Value::Bool(false));
                let current_done = self.emit_jump(Op::BR);

                self.patch_jump(not_done);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, current_key);
                let handled = self.emit_jump(Op::BR);

                self.patch_jump(need_start);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit(Op::GEN_NEXT);
                let has_more_slot = self.define_local("__php_gen_has_more");
                self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                let value_slot = self.define_local("__php_gen_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, started_key);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let no_more = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_generator_yield_value(value_slot);
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_generator_yield_value(value_slot);
                let start_done = self.emit_jump(Op::BR);

                self.patch_jump(no_more);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_u16(Op::STRUCT_SET, return_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_const(Value::Bool(false));

                self.patch_jump(current_done);
                self.patch_jump(handled);
                self.patch_jump(start_done);
            }
            "send" | "next" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let need_start = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let can_resume = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_const(Value::Bool(false));
                let done_already = self.emit_jump(Op::BR);

                self.patch_jump(can_resume);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                if field_name == "send" {
                    self.compile_expr(&arg_exprs[0])?;
                } else {
                    self.emit(Op::NULL);
                }
                self.emit_u16(Op::RESUME, 0);
                let value_slot = self.define_local("__php_gen_resume_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                self.emit_host_call(is_done_idx, 1);
                let yielded = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_u16(Op::STRUCT_SET, return_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let handled = self.emit_jump(Op::BR);

                self.patch_jump(yielded);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_generator_yield_value(value_slot);
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_generator_yield_value(value_slot);
                let resume_done = self.emit_jump(Op::BR);

                self.patch_jump(need_start);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit(Op::GEN_NEXT);
                let has_more_slot = self.define_local("__php_gen_has_more");
                self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                let start_value_slot = self.define_local("__php_gen_value");
                self.emit_u16(Op::LOCAL_SET, start_value_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, started_key);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let start_no_more = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_generator_yield_value(start_value_slot);
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_generator_yield_value(start_value_slot);
                let start_done = self.emit_jump(Op::BR);

                self.patch_jump(start_no_more);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::LOCAL_GET, start_value_slot);
                self.emit_u16(Op::STRUCT_SET, return_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, start_value_slot);

                self.patch_jump(done_already);
                self.patch_jump(handled);
                self.patch_jump(resume_done);
                self.patch_jump(start_done);
            }
            "throw" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, started_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let need_start = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, done_key);
                { let line = self.line; crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line); };
                let can_resume = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_const(Value::Bool(false));
                let done_already = self.emit_jump(Op::BR);

                self.patch_jump(can_resume);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(&arg_exprs[0])?;
                self.emit_generator_control_packet_from_stack("throw");
                self.emit_u16(Op::RESUME, 0);
                let value_slot = self.define_local("__php_gen_throw_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                self.emit_host_call(is_done_idx, 1);
                let yielded = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_u16(Op::STRUCT_SET, return_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let handled = self.emit_jump(Op::BR);

                self.patch_jump(yielded);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_generator_yield_value(value_slot);
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_generator_yield_value(value_slot);
                let resume_done = self.emit_jump(Op::BR);

                self.patch_jump(need_start);
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
                let start_no_more = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(&arg_exprs[0])?;
                self.emit_generator_control_packet_from_stack("throw");
                self.emit_u16(Op::RESUME, 0);
                let start_resume_slot = self.define_local("__php_gen_throw_resume_value");
                self.emit_u16(Op::LOCAL_SET, start_resume_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                self.emit_host_call(is_done_idx, 1);
                let start_yielded = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::LOCAL_GET, start_resume_slot);
                self.emit_u16(Op::STRUCT_SET, return_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, start_resume_slot);
                let start_handled = self.emit_jump(Op::BR);

                self.patch_jump(start_yielded);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_generator_yield_value(start_resume_slot);
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_generator_yield_value(start_resume_slot);
                let start_resume_done = self.emit_jump(Op::BR);

                self.patch_jump(start_no_more);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_u16(Op::STRUCT_SET, done_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::LOCAL_GET, start_value_slot);
                self.emit_u16(Op::STRUCT_SET, return_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::STRUCT_SET, current_key);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, start_value_slot);

                self.patch_jump(done_already);
                self.patch_jump(handled);
                self.patch_jump(resume_done);
                self.patch_jump(start_handled);
                self.patch_jump(start_resume_done);
            }
            _ => unreachable!(),
        }

        let end = self.emit_jump(Op::BR);
        self.patch_jump(not_gen);
        Ok(Some(end))
    }
}