//! Shared proxy/interception lowering helpers.
//!
//! ECMA owns the JS Proxy semantics and invariants in `platforms/ecma`.
//! This module is the common compiler boundary that chooses proxy operations
//! without making individual expression/statement paths know about language
//! hook storage.

use super::*;
use crate::primitives::instructions::recipes;

impl Compiler {
    fn proxy_hook_error(&self, op: &str) -> String {
        format!(
            "profile '{}' enabled proxy lowering but did not register proxy {} hook",
            self.profile.name, op
        )
    }

    /// Stack: `[target, handler] -> [proxy]`.
    pub(crate) fn emit_proxy_create(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_create
            .ok_or_else(|| self.proxy_hook_error("create"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object, key] -> [value]`.
    pub(crate) fn emit_proxy_get(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_get
            .ok_or_else(|| self.proxy_hook_error("get"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object, key, value] -> [value]`.
    pub(crate) fn emit_proxy_set(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_set
            .ok_or_else(|| self.proxy_hook_error("set"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Strict-mode set. Stack: `[object, key, value] -> [success_bool]`.
    pub(crate) fn emit_proxy_set_bool(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_set_bool
            .ok_or_else(|| self.proxy_hook_error("set_bool"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object, key] -> [bool]`.
    pub(crate) fn emit_proxy_has(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_has
            .ok_or_else(|| self.proxy_hook_error("has"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object, key] -> [bool]`.
    pub(crate) fn emit_proxy_delete_property(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_delete_property
            .ok_or_else(|| self.proxy_hook_error("delete_property"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[callable_or_proxy, this_arg, args_array] -> [result]`.
    pub(crate) fn emit_proxy_apply(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_apply
            .ok_or_else(|| self.proxy_hook_error("apply"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object] -> [keys_array]`.
    pub(crate) fn emit_proxy_own_keys(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_own_keys
            .ok_or_else(|| self.proxy_hook_error("own_keys"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object, key] -> [descriptor_or_undefined]`.
    pub(crate) fn emit_proxy_get_own_property_descriptor(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_get_own_property_descriptor
            .ok_or_else(|| self.proxy_hook_error("get_own_property_descriptor"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object, key, descriptor] -> [success_bool]`.
    pub(crate) fn emit_proxy_define_property(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_define_property
            .ok_or_else(|| self.proxy_hook_error("define_property"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object] -> [prototype_or_null]`.
    pub(crate) fn emit_proxy_get_prototype_of(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_get_prototype_of
            .ok_or_else(|| self.proxy_hook_error("get_prototype_of"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object, prototype_or_null] -> [success_bool]`.
    pub(crate) fn emit_proxy_set_prototype_of(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_set_prototype_of
            .ok_or_else(|| self.proxy_hook_error("set_prototype_of"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object] -> [bool]`.
    pub(crate) fn emit_proxy_is_extensible(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_is_extensible
            .ok_or_else(|| self.proxy_hook_error("is_extensible"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[object] -> [success_bool]`.
    pub(crate) fn emit_proxy_prevent_extensions(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_prevent_extensions
            .ok_or_else(|| self.proxy_hook_error("prevent_extensions"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Stack: `[constructor_or_proxy, args_array] -> [object]`.
    pub(crate) fn emit_proxy_construct(&mut self) -> Result<(), String> {
        let hook = vybe_runtime::registry::hooks(&self.profile.name)
            .proxy_construct
            .ok_or_else(|| self.proxy_hook_error("construct"))?;
        hook(&mut self.chunks, self.current, self.line);
        Ok(())
    }

    /// Shared attribute-miss interception role.
    ///
    /// Python `__getattr__` / `__getattribute__` and PHP `__get` publish the
    /// same protocol slot. Normal member reads stay ordinary ECMA `[[Get]]`
    /// first; only an `undefined` result probes the slot and calls the handler
    /// as `handler(receiver, name)`.
    ///
    /// Stack: `[value] -> [value]`.
    /// The WRITE twin of [`Self::emit_getattr_slot_probe`] — `ProtocolSlot::SetAttr`
    /// (PHP `__set`, Python `__setattr__`).
    ///
    /// Stack: `[value] -> []`. Returns `true` when it emitted the probe, so the
    /// caller knows the value was consumed and the ordinary store must be
    /// skipped; `false` leaves the stack untouched.
    ///
    /// `SetAttr` had NO reader: it sat in the slot vocabulary while every
    /// frontend that wanted a catch-all write synthesised a direct `__set`
    /// member call in its own walker. This is the consumer that makes the
    /// declaration mean something.
    pub(super) fn emit_setattr_slot_probe(&mut self, obj_slot: u16, field_name: &str) -> bool {
        if !self.program_has_setattr {
            return false;
        }
        let slot_key = vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::SetAttr);
        let line = self.line;
        let value_slot = self.define_local("__setattr_value");
        let handler_slot = self.define_local("__setattr_handler");

        self.emit_u16(Op::LOCAL_SET, value_slot);

        // Only a real object can carry the slot — `is_object` gates the probe
        // exactly as the read side does, because a `ecma:reflect.get` on a
        // primitive is not a lookup that can succeed.
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        recipes::is_object(self.chunk(), line);
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
        // `handler(receiver, name, value)` — the same receiver-first shape the
        // read probe uses, plus the value being written.
        self.emit_u16(Op::LOCAL_GET, handler_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(field_name)));
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::callable::emit_direct_invoke_chunk(self.chunk(), 3, line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        true
    }

    pub(super) fn emit_getattr_slot_probe(&mut self, obj_slot: u16, field_name: &str) {
        if !self.program_has_getattr {
            return;
        }
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

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        recipes::is_object(self.chunk(), line);
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
        self.emit_u16(Op::LOCAL_GET, handler_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(field_name)));
        crate::primitives::callable::emit_direct_invoke_chunk(self.chunk(), 2, line);
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.chunk().emit_end(line);

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, value_slot);
    }
}
