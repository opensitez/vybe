//! Shared proxy/interception lowering helpers.
//!
//! ECMA owns the JS Proxy semantics and invariants in `platforms/ecma`.
//! This module is the common compiler boundary that chooses proxy operations
//! without making individual expression/statement paths know about language
//! hook storage.

use super::Compiler;

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
    pub(crate) fn emit_proxy_delete_property(&mut self) {
        let idx = self.import("ecma:proxy", "deleteProperty");
        self.emit_host_call(idx, 2);
    }

    /// Stack: `[callable_or_proxy, this_arg, args_array] -> [result]`.
    pub(crate) fn emit_proxy_apply(&mut self) {
        let idx = self.import("ecma:proxy", "apply");
        self.emit_host_call(idx, 3);
    }
}
