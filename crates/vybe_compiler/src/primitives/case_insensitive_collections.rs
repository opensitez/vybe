//! Case-insensitive .NET collection calls — `Add`/`ContainsKey`/`Remove` on
//! `StringComparer.OrdinalIgnoreCase` dicts/sets.
//!
//! Enum handling → `primitives/enums.rs`; reflection → `primitives/reflection.rs`;
//! TryParse/TryGetValue are walker desugars.

use super::*;
use crate::primitives::calls::resolve_receiver_type_hint;

impl Compiler {
    pub(super) fn try_compile_dotnet_case_insensitive_collection_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        if !self.expr_uses_case_insensitive_string_keys(object) {
            return Ok(false);
        }

        let receiver_type = resolve_receiver_type_hint(self, object).unwrap_or_default();
        let normalized = Self::normalize_type_hint(&receiver_type);
        let line = self.line;

        if Self::is_dictionary_type_hint(&normalized) {
            match (field.as_str(), args.len()) {
                ("Add", 2) => {
                    let obj_slot = self.define_local("__dict_add_obj");
                    let key_slot = self.define_local("__dict_add_key");
                    let keys_slot = self.define_local("__dict_add_keys");

                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    self.compile_collection_key(object, &args[0].value)?;
                    self.emit_u16(Op::LOCAL_SET, key_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    self.compile_expr(&args[1].value)?;
                    let idx = self.import("ecma:map", "set");
                    self.emit_host_call(idx, 3);
                    self.emit(Op::DROP);

                    let keys_key = self.str_const("__keys");
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::STRUCT_GET, keys_key);
                    self.emit_u16(Op::LOCAL_SET, keys_slot);

                    self.emit_u16(Op::LOCAL_GET, keys_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_SET, keys_slot);
                    self.emit_u16(Op::STRUCT_SET, keys_key);
                    self.emit(Op::DROP);

                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_GET, keys_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                    return Ok(true);
                }
                ("ContainsKey", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:map", "has");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                ("Remove", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:map", "delete");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                _ => {}
            }
        }

        if normalized.contains("hashset") || normalized.contains("sortedset") {
            match (field.as_str(), args.len()) {
                ("Add", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    self.emit_common("dotnet.hashset_add", 2, line);
                    return Ok(true);
                }
                ("Contains", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:set", "has");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                ("Remove", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:set", "delete");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                _ => {}
            }
        }

        Ok(false)
    }
}
