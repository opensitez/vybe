//! Dynamic symbol resolution recipes.
//!
//! Most names resolve statically through the common namespace/class machinery.
//! Some languages also expose a runtime "missing symbol" hook: PHP class
//! autoload, Ruby constant missing, and potentially similar features later.
//! This module owns the shared bytecode shape for "try a symbol, invoke a
//! resolver, try again"; each language still owns its callback storage and
//! source spelling.

use crate::primitives::class_slots;
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

use super::*;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
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
    let slot = alloc_local(chunk);
    crate::primitives::globals::emit_read(chunk, stack_global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    crate::primitives::globals::emit_write(chunk, stack_global, line);
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
        crate::primitives::globals::emit_read(chunk, global, line);
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
            crate::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
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
            crate::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
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
    let symbol_slot = alloc_local(chunk);
    crate::primitives::globals::emit_read(chunk, global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);

    push_str(chunk, source_spelling, line);
    emit_resolver_stack_invoke(chunk, resolver, Some(global), line);

    crate::primitives::globals::emit_read(chunk, global, line);
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
    crate::primitives::globals::emit_read(chunk, primary_global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, symbol_slot, line);

    if let Some(fallback) = fallback_global {
        emit_fallback_if_undefined(chunk, symbol_slot, fallback, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    emit_undefined_test(chunk, line);
    chunk.emit_if(line);

    push_str(chunk, source_spelling, line);
    emit_resolver_stack_invoke(chunk, resolver, Some(primary_global), line);

    crate::primitives::globals::emit_read(chunk, primary_global, line);
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
    crate::primitives::globals::emit_read(chunk, fallback, line);
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

    // Not-undefined is not enough: a module global that was DECLARED but never
    // ASSIGNED reads `null`, which is exactly the shape of a conditionally
    // declared symbol whose branch did not run —
    // `if (false) { function f() {…} }` declares the global and never stores
    // the closure. Treating that as defined reported `function_exists('f')`
    // true for a function that never comes into being, which in turn made
    // `if (!function_exists('f')) { function f() {…} }` skip its own body.
    // Null is not a defined symbol in any language this serves.
    chunk.emit_op_u16(Op::LOCAL_GET, symbol_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);

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
    // `emit_dyn_eq` yields an i32, and every other arm of this function yields
    // a BOOL. Without this the answer's type depends on which arm ran:
    // `class_exists($c)` came back as integer `1`, so `=== true` was false and
    // `gettype()` said "integer". Only the stamp-present arm could reach it,
    // which is why a literal name looked correct.
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    class_slots::emit_class_alloc(chunk, line);
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
    crate::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], argc, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

impl Compiler {
    /// Resolve the symbol whose NAME is only known at runtime.
    ///
    /// Stack: `[] -> [symbol | undefined]`; the name expression is compiled
    /// here.
    ///
    /// Same shape, and for the same reason, as
    /// `globals::emit_global_namespace_index_get`: `GLOBAL_GET` takes the name
    /// as a u16 IMMEDIATE and the host surface is closed — no host fn may be
    /// added to reflect `vm.globals` — so a runtime key cannot index the map
    /// directly. What composes from existing ops is one comparison per
    /// DECLARED symbol, each guarding a real read. The read is
    /// [`Self::emit_constructor_global_ref`] rather than a raw global, so a
    /// name that matches a class still runs the language's autoloaders.
    ///
    /// A name the module never declared leaves `undefined`, which is exactly
    /// what a miss on the literal path reads, so both paths agree.
    ///
    /// ⚠ It answers for what the MODULE declared. A builtin that never becomes
    /// a global — php's `strlen` — is resolved by the frontend's own table
    /// when the name is a literal, and that table is not reachable from a
    /// runtime string. Names of that kind still miss.
    pub(crate) fn emit_symbol_ref_by_runtime_name(
        &mut self,
        name: &Expression,
    ) -> Result<(), String> {
        self.compile_expr(name)?;
        self.emit_symbol_ref_for_name_on_stack();
        Ok(())
    }

    /// Compile a **class designator** — an expression that denotes a class.
    ///
    /// A designator is either the class itself or a STRING naming it, and
    /// which one it is cannot be known before the expression runs. So the
    /// question is asked at run time rather than decided by the frontend:
    /// `typeof x === "string"` resolves the name, anything else is already the
    /// class. PHP states the same rule in its own error — *"Class name must be
    /// a valid object or a string"* — and it is expressed here without a
    /// profile flag, a language name or a per-language table, so every
    /// frontend that has designators gets it.
    ///
    /// `New` and `StaticAccess` are the two nodes that carry one. Both call
    /// this, so `new $cls()` and `$cls::method()` cannot answer differently.
    ///
    /// Stack: `[] -> [class]`.
    pub(crate) fn emit_class_designator(&mut self, class: &Expression) -> Result<(), String> {
        let line = self.line;
        // Both slots are defined BEFORE the branch. `emit_symbol_ref_for_name_on_stack`
        // opens one `if`/`end` per declared symbol, and nesting that inside a
        // value-producing block while also defining its locals there emitted
        // bytecode the VM rejected outright (`Invalid opcode`). A plain `if`
        // that writes a pre-declared slot keeps every block value-less, so the
        // nesting depth stops mattering.
        let slot = self.define_local("__class_designator");
        let out = self.define_local("__class_designator_out");
        self.compile_expr(class)?;
        self.emit_u16(Op::LOCAL_SET, slot);
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_u16(Op::LOCAL_SET, out);

        self.emit_u16(Op::LOCAL_GET, slot);
        crate::primitives::reflection::emit_typeof_in_chunk(self.chunk(), line);
        self.emit_const(Value::String(std::sync::Arc::from("string")));
        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_symbol_ref_for_name_on_stack();
        self.emit_u16(Op::LOCAL_SET, out);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, out);
        Ok(())
    }

    /// Resolve a **callable designator** held in `slot`, in place.
    ///
    /// The sibling of [`Self::emit_class_designator`], asking the same kind of
    /// question one step further out: not *which class does this value denote*
    /// but *which callable*. Which spelling a value is, is a property of the
    /// VALUE, so it is settled at run time.
    ///
    /// One spelling is resolved here: a two-element `[receiver, "method"]`
    /// pair whose first element is an OBJECT. The result is the method BOUND
    /// to that receiver — one callable value, which is what lets every asker
    /// (a call, `is_callable`, a callback handed to a sort) read the same
    /// answer with no receiver out-parameter.
    ///
    /// A plain function value, a string naming a declared function, and an
    /// object filling [`ProtocolSlot::Call`](vybe_ast::ProtocolSlot::Call) are
    /// already answered — by the call site, by
    /// [`Self::emit_source_function_callable_name_resolution`] and by the Call
    /// slot probe respectively.
    ///
    /// ⛔ A class NAME that is only known at run time — `"Class::method"`, or a
    /// pair whose first element is a string — is deliberately NOT resolved
    /// here, and the reason is a hard size limit rather than a missing idea.
    /// Resolving a name needs [`Self::emit_symbol_ref_for_name_on_stack`], one
    /// comparison per declared symbol each guarding the language's autoload
    /// sequence. That is affordable at a `New` or `StaticAccess` site, which
    /// is why `emit_class_designator` can inline it; it is not affordable at
    /// EVERY call site. `Chunk::emit_try_table_clauses` patches a catch
    /// handler as a two-byte offset, so one inflated call site inside a `try`
    /// pushed the body past 65535 bytes, the offset truncated, and execution
    /// resumed mid-instruction (`Invalid opcode: 0x0000 0x2000`). Those two
    /// spellings are resolved by the frontend instead, where the name is
    /// literal — see `php_callable_target_expr`.
    ///
    /// Gated on `source_function_callable_aliases`, the axis that already
    /// declares *"this language has PHP/Ruby-style dynamic callables"*. It has
    /// to be gated on something: the VM reads an ARRAY callee as a .NET
    /// multicast delegate, so a language that means that must not have its
    /// arrays claimed here.
    pub(crate) fn emit_callable_designator_in_slot(&mut self, slot: u16) {
        if !self.profile.source_function_callable_aliases {
            return;
        }
        let line = self.line;

        // Every local is declared BEFORE the first branch — see the note in
        // `emit_class_designator`.
        let member_slot = self.define_local("__callable_designator_member");
        let receiver_slot = self.define_local("__callable_designator_receiver");
        let resolved_slot = self.define_local("__callable_designator_resolved");

        // Guarded by `isArray`, not by a length probe on any object:
        // `collections::emit_len` falls back to `Object.keys().length`, so an
        // ordinary two-field instance would have answered 2 and been taken
        // apart as a pair.
        self.emit_u16(Op::LOCAL_GET, slot);
        let is_array_idx = self.import("ecma:array", "isArray");
        self.emit_host_call(is_array_idx, 1);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, slot);
        let length_idx = self.import("ecma:array", "length");
        self.emit_host_call(length_idx, 1);
        self.emit_const(Value::I32(2));
        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_const(Value::I32(1));
        self.emit(Op::ARRAY_GET);
        self.emit_u16(Op::LOCAL_SET, member_slot);

        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_const(Value::I32(0));
        self.emit(Op::ARRAY_GET);
        self.emit_u16(Op::LOCAL_SET, receiver_slot);

        // The first element must be an OBJECT to be a receiver. A STRING
        // there names a CLASS, which this cannot resolve — see the note
        // above — and anything else is not a callable pair at all: `[1, 2]`
        // is just an array, and reading member `2` off the number `1` threw
        // where php answers `false`.
        self.emit_u16(Op::LOCAL_GET, receiver_slot);
        crate::primitives::reflection::emit_typeof_in_chunk(self.chunk(), line);
        self.emit_const(Value::String(std::sync::Arc::from("object")));
        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        // `ecma:function.bind` supplies the fresh function object — mutating
        // the method read off the instance is not an option, since a class
        // whose methods are not distinct per instance hands back the INTERNED
        // funcref and the write would land on every instance at once.
        //
        // The receiver is then stamped as `__vybe_method_receiver`, which is
        // the property every call path in this compiler reads to pass the
        // receiver as the leading ARGUMENT. `bind` alone sets the ECMA `this`,
        // and a method compiled with an explicit receiver parameter never
        // reads that: the first real argument landed in the receiver's slot
        // and `hi($n)` printed an empty `$n`.
        self.emit_u16(Op::LOCAL_GET, receiver_slot);
        self.emit_u16(Op::LOCAL_GET, member_slot);
        crate::primitives::reflection::emit_get_property_in_chunk(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, resolved_slot);

        // Only a member that IS there gets bound. `[1, 2]` is a two-element
        // array and not a callable: reading member `2` off `1` yields nothing,
        // and binding nothing threw where php just answers `false`.
        self.emit_u16(Op::LOCAL_GET, resolved_slot);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, resolved_slot);
        self.emit_u16(Op::LOCAL_GET, receiver_slot);
        let bind_idx = self.import("ecma:function", "bind");
        self.emit_host_call(bind_idx, 2);
        self.emit_u16(Op::LOCAL_SET, slot);
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_u16(Op::LOCAL_GET, receiver_slot);
        self.class_set(
            class_slots::ObjSource::Stack,
            &class_slots::ClassSlot::internal("__vybe_method_receiver"),
            class_slots::ValueSource::Stack,
        );

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
    }

    /// [`Self::emit_symbol_ref_by_runtime_name`] with the name already on the
    /// stack. Stack: `[name] -> [symbol | undefined]`.
    pub(crate) fn emit_symbol_ref_for_name_on_stack(&mut self) {
        let line = self.line;
        let key_slot = self.define_local("__symbol_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);

        let result_slot = self.define_local("__symbol_result");
        crate::primitives::instructions::core_wasm::undefined(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, result_slot);

        let mut candidates: Vec<String> = self.defined_globals.iter().cloned().collect();
        candidates.sort();
        for global in candidates {
            self.emit_u16(Op::LOCAL_GET, key_slot);
            self.emit_const(Value::String(std::sync::Arc::from(global.as_str())));
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_constructor_global_ref(&global, &global);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, result_slot);
    }

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
            self.emit_global_read(ctor_global);
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
            self.emit_global_read(primary_ctor_global);
        }
    }
}
