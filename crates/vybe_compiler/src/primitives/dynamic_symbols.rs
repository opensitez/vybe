//! Dynamic symbol resolution recipes.
//!
//! Most names resolve statically through the common namespace/class machinery.
//! Some languages also expose a runtime "missing symbol" hook: PHP class
//! autoload, Ruby constant missing, and potentially similar features later.
//! This module owns the shared bytecode shape for "try a symbol, invoke a
//! resolver, try again"; each language still owns its callback storage and
//! source spelling.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

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

/// A registered stack of symbol resolvers.
///
/// One global holds an ordered array of resolver callables, consulted in turn
/// when a symbol misses. This module owns the bytecode shape; the language owns
/// every spelling — the global's name, the source-level registration functions,
/// and its "callable object" protocol member.
#[derive(Clone, Copy)]
pub struct ResolverStack<'a> {
    /// Global holding the ordered array of resolver callables.
    pub stack_global: &'a str,
    /// Member consulted when a stack entry is an object rather than a plain
    /// function — the language's callable-object protocol member (PHP
    /// `__invoke`, Python `__call__`). `None` accepts functions only.
    pub invoke_member: Option<&'a str>,
}

fn array_call(chunk: &mut Chunk, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import("ecma:array", name);
    chunk.emit_call(idx, argc, line);
}

/// Push the resolver array, creating and storing an empty one when the global
/// is still undefined. Stack on exit: `[array]`.
fn emit_stack_load(chunk: &mut Chunk, stack_global: &str, line: u32) {
    let global_idx = str_idx(chunk, stack_global);
    let slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::GLOBAL_GET, global_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, global_idx, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

/// Add a resolver to the stack. Stack on entry: `[callable, prepend_flag]`;
/// exit: `[bool]`.
///
/// Re-registering a callable already on the stack is a no-op, so a resolver
/// never runs twice for one symbol. Entries are stored exactly as supplied —
/// identity is what [`emit_resolver_unregister`] matches on, so wrapping them
/// here would silently break removal.
pub fn emit_resolver_register(chunk: &mut Chunk, stack: ResolverStack<'_>, line: u32) {
    let prepend_slot = alloc_local(chunk);
    let callback_slot = alloc_local(chunk);
    let array_slot = alloc_local(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, prepend_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, callback_slot, line);
    emit_stack_load(chunk, stack.stack_global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, array_slot, line);

    // if (indexOf(stack, callable) < 0)
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    array_call(chunk, "indexOf", 2, line);
    chunk.emit_i32_const(0, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, prepend_slot, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    array_call(chunk, "unshift", 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    array_call(chunk, "push", 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_bool_const(true, line);
}

/// Remove a resolver from the stack by identity. Stack on entry: `[callable]`;
/// exit: `[bool]` — whether it was present.
pub fn emit_resolver_unregister(chunk: &mut Chunk, stack: ResolverStack<'_>, line: u32) {
    let callback_slot = alloc_local(chunk);
    let array_slot = alloc_local(chunk);
    let found_slot = alloc_local(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, callback_slot, line);
    emit_stack_load(chunk, stack.stack_global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, array_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    array_call(chunk, "indexOf", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, found_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, found_slot, line);
    chunk.emit_i32_const(0, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, found_slot, line);
    chunk.emit_i32_const(1, line);
    array_call(chunk, "splice", 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// A copy of the registered resolvers, in call order. Stack on exit:
/// `[array]`.
pub fn emit_resolver_list(chunk: &mut Chunk, stack: ResolverStack<'_>, line: u32) {
    // slice(stack, 0) copies through to the end, so callers can't mutate the
    // live queue through the returned value.
    emit_stack_load(chunk, stack.stack_global, line);
    chunk.emit_i32_const(0, line);
    array_call(chunk, "slice", 2, line);
}

/// Run resolvers in order against a symbol name, stopping as soon as
/// `resolved_global` becomes defined. Stack on entry: `[name]`; exit: `[]`.
///
/// `resolved_global` is `None` when the caller cannot name the global the
/// resolvers are expected to define (a computed symbol name); every resolver
/// then runs.
pub fn emit_resolver_stack_invoke(
    chunk: &mut Chunk,
    stack: ResolverStack<'_>,
    resolved_global: Option<&str>,
    line: u32,
) {
    let name_slot = alloc_local(chunk);
    let array_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let index_slot = alloc_local(chunk);
    let entry_slot = alloc_local(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    emit_stack_load(chunk, stack.stack_global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, array_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    array_call(chunk, "length", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer = chunk.emit_block(line);
    let (body, _) = chunk.emit_loop_s(line);

    // while (index < len)
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    // Already resolved by an earlier resolver — stop.
    if let Some(global) = resolved_global {
        let idx = str_idx(chunk, global);
        chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
        emit_undefined_test(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_br_if(1, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    array_call(chunk, "get", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entry_slot, line);

    // A plain function is called directly; a callable object goes through the
    // language's protocol member.
    match stack.invoke_member {
        Some(member) => {
            chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
            let type_of = chunk.add_import("ecma:value", "typeof");
            chunk.emit_call(type_of, 1, line);
            chunk.emit_string_const("function", line);
            crate::primitives::ops::emit_dyn_eq(chunk, line);
            crate::primitives::ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_if(line);
            chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
            chunk.emit_op_u8(Op::CALL_REF, 1, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_else(line);
            // Through `emit_invoke_method`, not a raw `ecma:value.invokeMethod`
            // import: it also binds (and restores) `__js_this`, which the
            // resolver body needs to see its own receiver.
            chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
            crate::primitives::invoke::emit_invoke_method(
                std::slice::from_mut(chunk),
                0,
                member,
                1,
                line,
            );
            chunk.emit_op(Op::DROP, line);
            chunk.emit_end(line);
        }
        None => {
            chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
            chunk.emit_op_u8(Op::CALL_REF, 1, line);
            chunk.emit_op(Op::DROP, line);
        }
    }

    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_i32_const(1, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(body);
    chunk.emit_end(line);
    chunk.patch_block(outer);
}

/// Push a reference to `global`, consulting the resolver stack with
/// `source_spelling` if the global is currently undefined. Stack on exit:
/// `[symbol_ref]`.
pub fn emit_registered_global_ref(
    chunk: &mut Chunk,
    global: &str,
    source_spelling: &str,
    resolver: ResolverStack<'_>,
    line: u32,
) {
    let global_idx = str_idx(chunk, global);
    let symbol_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::GLOBAL_GET, global_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);

    push_str(chunk, source_spelling, line);
    emit_resolver_stack_invoke(chunk, resolver, Some(global), line);

    chunk.emit_op_u16(Op::GLOBAL_GET, global_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
}

/// Like [`emit_registered_global_ref`], but checks an optional fallback global
/// before and after consulting the resolver stack.
pub fn emit_registered_dynamic_global_ref(
    chunk: &mut Chunk,
    primary_global: &str,
    fallback_global: Option<&str>,
    source_spelling: &str,
    resolver: ResolverStack<'_>,
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

    push_str(chunk, source_spelling, line);
    emit_resolver_stack_invoke(chunk, resolver, Some(primary_global), line);

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

/// Whether a declared symbol resolves — and, when `expected_kind` is given,
/// whether it was declared as that kind.
///
/// This is the one primitive behind the whole `*_exists` family:
/// `class_exists` / `interface_exists` / `trait_exists` / `enum_exists` differ
/// only in which `ReflectKind` they accept. The kind comes from the annotation
/// the class compiler stamps (`reflection::FIELD_KIND`), so it is answered from
/// the runtime object rather than a compile-time per-language table — which is
/// what lets it be true for a type an autoloader defined after compilation.
///
/// Stack: `[symbol_ref] -> [bool]`. The caller decides whether to resolve the
/// reference through the resolver stack first, which is how the language's
/// "autoload" flag is honoured.
pub fn emit_symbol_kind_test(
    chunk: &mut Chunk,
    expected_kind: Option<crate::primitives::reflection::ReflectKind>,
    line: u32,
) {
    let symbol_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);

    let Some(kind) = expected_kind else {
        // No kind constraint: defined is the whole answer.
        crate::primitives::ops::emit_i32_to_bool(chunk, line);
        return;
    };

    // Defined, and the kind does not CONTRADICT what was asked.
    //
    // The annotation is only carried by types that went through the shared
    // class compiler. Host- and prelude-provided types (PHP `DateTime`,
    // `ArrayObject`, SPL, PDO) have a real constructor global and no stamp, so
    // requiring `__kind == kind` reports them missing — and
    // `if (class_exists('PDO'))` is the standard feature-detection idiom.
    // Absent stamp therefore falls back to definedness; only a stamp that is
    // PRESENT AND DIFFERENT is a rejection. That still discriminates every
    // type the compiler declared, which is where the question is meaningful.
    let kind_slot = alloc_local(chunk);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    chunk.emit_string_const(crate::primitives::reflection::FIELD_KIND, line);
    let reflect_get = chunk.add_import("ecma:reflect", "get");
    chunk.emit_call(reflect_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, kind_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunk.emit_string_const(kind.as_str(), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// Throw `exception_name(message)` when the value on the stack is undefined,
/// otherwise leave it there.
///
/// This is the "resolve or fail" half of dynamic symbol lookup: Java
/// `Class.forName` raises `ClassNotFoundException`, PHP raises `Error`, Python
/// and Ruby raise `NameError`. The mechanism is identical in all of them, so
/// only the exception's spelling differs — and that arrives as data from the
/// language's profile, never as a check in here. `canonical_exception_name`
/// then normalizes it, so a Java `ClassNotFoundException` stays catchable
/// across the language boundary.
///
/// Stack: `[symbol_ref] -> [symbol_ref]`, or throws.
pub fn emit_throw_if_unresolved(chunk: &mut Chunk, exception_name: &str, message: &str, line: u32) {
    let symbol_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    push_str(chunk, message, line);
    crate::primitives::errors::emit_exception_new_finalize(chunk, exception_name, line);
    crate::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
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
            vybe_runtime::registry::hooks(&self.profile.name)
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
            vybe_runtime::registry::hooks(&self.profile.name)
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
