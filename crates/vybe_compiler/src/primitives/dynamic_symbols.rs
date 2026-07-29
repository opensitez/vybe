//! Dynamic symbol resolution recipes.
//!
//! Most names resolve statically through the common namespace/class machinery.
//! Some languages also expose a runtime "missing symbol" hook: PHP class
//! autoload, Ruby constant missing, and potentially similar features later.
//! This module owns the shared bytecode shape for "try a symbol, invoke a
//! resolver, try again"; each language still owns its callback storage and
//! source spelling.

use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use super::*;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn str_idx(chunk: &mut Chunk, value: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(value)))
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
}

fn emit_undefined_test(chunk: &mut Chunk, line: u32) {
    let undef_test = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef_test, 1, line);
}

/// A registered language resolver stored in globals.
pub struct RegisteredResolver<'a> {
    pub callback_global: &'a str,
    pub receiver_global: &'a str,
}

/// Push a reference to `global`, invoking `resolver` with `source_spelling` if
/// the global is currently undefined. Stack on exit: `[symbol_ref]`.
pub fn emit_registered_global_ref(
    chunk: &mut Chunk,
    global: &str,
    source_spelling: &str,
    resolver: RegisteredResolver<'_>,
    line: u32,
) {
    let global_idx = str_idx(chunk, global);
    let symbol_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::GLOBAL_GET, global_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);

    emit_registered_resolver_invoke(chunk, source_spelling, resolver, line);

    chunk.emit_op_u16(Op::GLOBAL_GET, global_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
}

/// Like [`emit_registered_global_ref`], but checks an optional fallback global
/// before and after the resolver invocation.
pub fn emit_registered_dynamic_global_ref(
    chunk: &mut Chunk,
    primary_global: &str,
    fallback_global: Option<&str>,
    source_spelling: &str,
    resolver: RegisteredResolver<'_>,
    line: u32,
) {
    let symbol_slot = alloc_local(chunk);
    let primary_idx = str_idx(chunk, primary_global);
    chunk.emit_op_u16(Op::GLOBAL_GET, primary_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);

    if let Some(fallback) = fallback_global {
        emit_fallback_if_undefined(chunk, symbol_slot, fallback, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);

    emit_registered_resolver_invoke(chunk, source_spelling, resolver, line);

    chunk.emit_op_u16(Op::GLOBAL_GET, primary_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);
    if let Some(fallback) = fallback_global {
        emit_fallback_if_undefined(chunk, symbol_slot, fallback, line);
    }

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
}

fn emit_fallback_if_undefined(chunk: &mut Chunk, symbol_slot: u16, fallback: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);
    let fallback_idx = str_idx(chunk, fallback);
    chunk.emit_op_u16(Op::GLOBAL_GET, fallback_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);
    chunk.emit_end(line);
}

/// Invoke a resolver callback saved in globals. When the receiver global is
/// undefined, the callback is invoked as a plain function with the symbol name.
/// Otherwise it is invoked with `(receiver, symbol_name)`.
pub fn emit_registered_resolver_invoke(
    chunk: &mut Chunk,
    source_spelling: &str,
    resolver: RegisteredResolver<'_>,
    line: u32,
) {
    let callback_slot = alloc_local(chunk);
    let callback_idx = str_idx(chunk, resolver.callback_global);
    chunk.emit_op_u16(Op::GLOBAL_GET, callback_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, callback_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);

    let receiver_slot = alloc_local(chunk);
    let receiver_idx = str_idx(chunk, resolver.receiver_global);
    chunk.emit_op_u16(Op::GLOBAL_GET, receiver_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    push_str(chunk, source_spelling, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    push_str(chunk, source_spelling, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Ruby-style receiver-local missing constant dispatch. If `target[name]`
/// misses, read `target[resolver_member]` and call it with `(target, name)`.
/// Stack on exit: `[constant_value_or_null]`.
pub fn emit_receiver_missing_symbol_get(
    chunks: &mut [Chunk],
    current: usize,
    target_slot: u16,
    name_slot: u16,
    resolver_member: &str,
    include_receiver_arg: bool,
    line: u32,
) {
    let resolver_slot = chunks[current].alloc_scratch(1);
    let reflect_get = chunks[current].add_import("ecma:reflect", "get");
    chunks[current].emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunks[current].emit_string_const(resolver_member, line);
    chunks[current].emit_call(reflect_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, resolver_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, resolver_slot, line);
    emit_undefined_test(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, resolver_slot, line);
    if include_receiver_arg {
        chunks[current].emit_op_u16(Op::LOCAL_GET, target_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    let argc = if include_receiver_arg { 2 } else { 1 };
    chunks[current].emit_op_u8(Op::CALL_REF, argc, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

impl Compiler {
    /// Push a reference to the class constructor global `ctor_global`.
    /// Dynamic-symbol-aware languages can invoke their registered type
    /// resolver before the final lookup; others use a plain global read.
    pub(crate) fn emit_constructor_global_ref(&mut self, ctor_global: &str, source_name: &str) {
        if self.profile.supports_autoload {
            let line = self.line;
            vybe_bytecode::registry::hooks(&self.profile.name)
                .constructor_ref_autoload
                .unwrap()(self.chunk(), ctor_global, source_name, line);
        } else {
            let idx = self.str_const(ctor_global);
            self.emit_u16(Op::GLOBAL_GET, idx);
        }
    }

    /// Like [`Self::emit_constructor_global_ref`] but resolves a primary
    /// constructor global then an optional fallback before invoking the dynamic
    /// type resolver.
    pub(crate) fn emit_dynamic_constructor_global_ref(
        &mut self,
        primary_ctor_global: &str,
        fallback_ctor_global: Option<&str>,
        source_name: &str,
    ) {
        if self.profile.supports_autoload {
            let line = self.line;
            vybe_bytecode::registry::hooks(&self.profile.name)
                .dynamic_constructor_ref_autoload
                .unwrap()(
                self.chunk(),
                primary_ctor_global,
                fallback_ctor_global,
                source_name,
                line,
            );
        } else {
            let idx = self.str_const(primary_ctor_global);
            self.emit_u16(Op::GLOBAL_GET, idx);
        }
    }
}
