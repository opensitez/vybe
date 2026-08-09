//! Function compilation helpers — shared bytecode patterns for function scaffolding.
//!
//! Every language compiles functions the same way at the bytecode level:
//! - Create a Chunk (name, arity)
//! - Map params to local slots
//! - Handle default values
//! - Compile body (language-specific)
//! - Emit null + return as safety net
//! - Store ref_func as local/global
//!
//! The scaffolding (everything except body compilation) is identical.
//! Python `def`, Dart `void f()`, JS `function`, C# `void F()`, VB `Sub`
//! all produce the same Chunk structure.

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::Compiler;

// ── Default parameter handling ──────────────────────────────────────────

/// Emit the start of a default parameter check.
/// If the parameter at `param_slot` is null (missing arg), the caller should compile
/// the default expression, then call `emit_default_param_end`.
/// Returns a structured block patch to close.
/// Stack: unchanged
pub fn emit_default_param_start(chunk: &mut Chunk, param_slot: u16, line: u32) -> usize {
    chunk.emit_op_u16(Op::LOCAL_GET, param_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    block
}

/// Emit the end of a default parameter check.
/// Caller must have compiled the default expression onto the stack.
/// Stack before: [default_value]  Stack after: [] (stored in param_slot)
pub fn emit_default_param_end(chunk: &mut Chunk, param_slot: u16, block: usize, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, param_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(block);
}

// ── Function chunk scaffolding ──────────────────────────────────────────

/// Create a new function chunk with the given name and arity.
/// Returns the chunk — caller adds it to their chunks vec and manages the scope.
pub fn create_function_chunk(name: &str, arity: u8) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = arity;
    chunk
}

/// Emit the function epilogue: null return (safety net for functions that
/// fall through without explicit return).
/// Stack: [] → diverges (return)
pub fn emit_function_epilogue(chunk: &mut Chunk, line: u32) {
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);
}

/// Start shared async-body scaffolding for function-like chunks.
/// Returns the catch jump to patch once the body has been emitted.
/// The caller remains responsible for Promise.resolve / Promise.reject
/// wrapping because import indices live at the compiler layer.
pub fn emit_async_body_start(chunk: &mut Chunk, line: u32) -> usize {
    crate::primitives::errors::emit_try_start(chunk, line)
}

/// Finish the normal fallthrough path of a shared async body.
/// Leaves `undefined` on the stack so the compiler can wrap it with
/// `Promise.resolve(undefined)` before returning.
pub fn emit_async_body_fallthrough(chunk: &mut Chunk, catch_jump: usize, line: u32) {
    crate::primitives::errors::emit_try_end(chunk, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    let _ = catch_jump;
}

/// Patch the catch edge for a shared async body so the compiler can
/// emit its rejection path.
pub fn patch_async_body_catch(chunk: &mut Chunk, catch_jump: usize) {
    crate::primitives::errors::patch_catch(chunk, catch_jump);
}

/// Emit ref_func to push a closure reference onto the stack.
/// `func_chunk_idx`: the chunk index of the compiled function.
/// `upvalue_count`: 0 for most functions, >0 for closures.
/// Stack: [] → [closure_ref]
pub fn emit_ref_func(chunk: &mut Chunk, func_chunk_idx: usize, upvalue_count: u8, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_chunk_idx as u16, line);
    chunk.emit(upvalue_count, line);
}

/// One closure upvalue descriptor following `emit_ref_func`:
/// `is_local` flag byte + u16 index (big-endian, like every other u16
/// operand). For `is_local` the index is the PARENT LOCAL SLOT (u16 to
/// match `Local.slot` — a u8 here once truncated slot 1321 to 41 and
/// silently captured garbage); otherwise it is the parent's upvalue
/// list position.
pub fn emit_closure_upvalue(chunk: &mut Chunk, is_local: bool, index: u16, line: u32) {
    chunk.emit(if is_local { 1 } else { 0 }, line);
    chunk.emit((index >> 8) as u8, line);
    chunk.emit((index & 0xff) as u8, line);
}

/// Store a function as a global variable.
/// Caller must have closure_ref on stack (from emit_ref_func).
/// Stack before: [closure_ref]  Stack after: []
pub fn emit_store_global_func(chunk: &mut Chunk, name: &str, line: u32) {
    crate::primitives::globals::emit_write(chunk, name, line);
}

/// Store a function in a local slot.
/// Caller must have closure_ref on stack (from emit_ref_func).
/// Stack before: [closure_ref]  Stack after: []
pub fn emit_store_local_func(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

// ── Cross-language function call ────────────────────────────────────────

/// Emit a call to a function by global name.
/// Pushes the function ref, then caller pushes args, then calls emit_call_args.
/// Stack: [] → [function_ref]
pub fn emit_push_global_func(chunk: &mut Chunk, name: &str, line: u32) {
    crate::primitives::globals::emit_read(chunk, name, line);
}

impl Compiler {
    pub(crate) fn source_function_callable_global_name_for_canon(
        &self,
        canon_name: &str,
    ) -> Option<String> {
        (self.profile.source_function_callable_aliases
            && self.defined_functions.contains(canon_name))
        .then(|| format!("__vybe_func${canon_name}"))
    }

    pub(crate) fn source_function_callable_global_name(&self, name: &str) -> Option<String> {
        let canon_name = self.canon(name);
        self.source_function_callable_global_name_for_canon(&canon_name)
    }

    /// The alias global a source-declared function's callable is assigned to,
    /// WITHOUT asking whether the declaration has been walked past yet.
    ///
    /// `defined_functions` holds only the declarations the emitter has already
    /// reached, so gating on it makes the alias name depend on WHERE in the
    /// file it is asked for. The call path can live with that — a call above a
    /// declaration compiles against the plain global, which module installation
    /// publishes up front, and that is exactly how a hoisted forward call
    /// works. A question ABOUT the declaration cannot: `function_exists('f')`
    /// above the declaration and the identical call below it compiled against
    /// two different globals and answered differently (measured: `true` then
    /// `false` for the same never-declared `f`).
    ///
    /// The alias name is a pure function of the canonical name, so compute it
    /// as one. `None` means the profile has no alias convention at all.
    /// The separator is normalized on top of `canon`: a declaration in
    /// `namespace Local` publishes `__vybe_func$Local.active`, but the name as
    /// WRITTEN in `function_exists('Local\active')` keeps its backslash through
    /// `canon`, so the question read `__vybe_func$Local\active` and answered
    /// false for a function that exists (measured in `-d`: `global.set
    /// __vybe_func$Local.active` against `global.get __vybe_func$Local\active`).
    /// Case is left alone — the alias is published in the declaration's own
    /// case, and the runtime-name path folds case on the needle instead.
    pub(crate) fn source_function_callable_alias_name(&self, name: &str) -> Option<String> {
        self.profile.source_function_callable_aliases.then(|| {
            let canon_name = self.canon(name);
            format!(
                "__vybe_func${}",
                crate::primitives::namespaces::normalize_source_path(&canon_name)
            )
        })
    }

    /// Whether a source-declared function exists, asked with a name that is
    /// only known at RUNTIME — the dynamic counterpart of the literal
    /// `function_exists('f')` arm. Stack: `[] -> [bool]`, name read from
    /// `name_slot` (the caller has already established it is a string).
    ///
    /// Same corpus and same comparison chain as
    /// [`Self::emit_source_function_callable_name_resolution`], the call path
    /// behind `$f = 'greet'; $f();` — existence has to agree with callability,
    /// so both answer from the same place. Only the payload differs: the call
    /// path swaps the string for the callable, this one resolves to the
    /// callable and then asks the runtime whether it is actually there.
    ///
    /// Answering `true` on a name MATCH alone — what this used to do — is the
    /// same bug the literal arm had. A match means the compiler SAW a
    /// declaration, not that the declaration ran, so
    /// `if (false) { function f() {…} }` reported `f` as existing.
    ///
    /// The miss sentinel is NULL, not the name string.
    /// `emit_symbol_kind_test` falls back to plain definedness when the
    /// `__kind` stamp is absent, so leaving the unmatched name on the stack
    /// would answer `true` for every string.
    pub(crate) fn emit_source_function_exists_by_runtime_name(&mut self, name_slot: u16) {
        let resolved_slot = self.define_local("__function_exists_resolved");
        self.emit_null();
        self.emit_u16(Op::LOCAL_SET, resolved_slot);

        let mut known_functions: Vec<String> = self.defined_functions.iter().cloned().collect();
        known_functions.sort();

        if self.profile.source_function_callable_aliases && !known_functions.is_empty() {
            self.emit_u16(Op::LOCAL_GET, name_slot);
            let line = self.line;
            crate::primitives::common::strings::emit_to_lower(self.chunk(), line);
            // Normalize the NEEDLE once and compare against the canonical
            // corpus — the same trade the call path makes.
            crate::primitives::namespaces::emit_normalize_source_path(self.chunk(), line);
            let needle_slot = self.define_local("__function_exists_needle");
            self.emit_u16(Op::LOCAL_SET, needle_slot);

            let matched_slot = self.define_local("__function_exists_matched");
            self.emit_const(Value::I32(0));
            self.emit_u16(Op::LOCAL_SET, matched_slot);

            for function_name in known_functions {
                let canonical = crate::primitives::namespaces::normalize_source_path(
                    &function_name.to_ascii_lowercase(),
                );
                let rooted = crate::primitives::namespaces::rooted_lookup_key(&canonical);
                for lowered_name in [canonical, rooted] {
                    self.emit_u16(Op::LOCAL_GET, matched_slot);
                    self.emit(Op::I32_EQZ);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    self.emit_raw_string_slot_eq_literal(needle_slot, lowered_name.as_str());
                    let line = self.line;
                    self.chunk().emit_if(line);

                    if let Some(callable_global) =
                        self.source_function_callable_global_name_for_canon(&function_name)
                    {
                        self.emit_global_read(&callable_global);
                        self.emit_u16(Op::LOCAL_SET, resolved_slot);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, matched_slot);
                    }
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                }
            }
        }

        // ONE kind test, after the chain — it allocates locals and emits an
        // `ecma:reflect` read, so per-candidate would scale the emitted code
        // with the function count for no extra answer.
        self.emit_u16(Op::LOCAL_GET, resolved_slot);
        let line = self.line;
        crate::primitives::dynamic_symbols::emit_symbol_kind_test(
            self.chunk(),
            Some(crate::primitives::reflection::ReflectKind::Function),
            line,
        );
    }

    pub(crate) fn emit_source_function_callable_name_resolution(&mut self, callee_slot: u16) {
        if !self.profile.source_function_callable_aliases {
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
            crate::primitives::instructions::host::CapabilityContext::get()
                .functions
                .emit(&mut self.chunks[self.current], "ecma:value", "typeof", 1, l);
        };
        self.emit_string_eq_literal("string");
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, callee_slot);
        let line = self.line;
        crate::primitives::common::strings::emit_to_lower(self.chunk(), line);
        // Normalize the NEEDLE once, then compare against the canonical corpus.
        // The alternative — expanding every known function back into every
        // separator spelling — is the same work done once per known function.
        crate::primitives::namespaces::emit_normalize_source_path(self.chunk(), line);
        let callee_name_slot = self.define_local("__source_string_callee_name");
        self.emit_u16(Op::LOCAL_SET, callee_name_slot);

        let matched_slot = self.define_local("__source_string_callee_matched");
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        for function_name in known_functions {
            // Lowercased to match the runtime needle, which `emit_to_lower`
            // has already folded.
            let canonical = crate::primitives::namespaces::normalize_source_path(
                &function_name.to_ascii_lowercase(),
            );
            let rooted = crate::primitives::namespaces::rooted_lookup_key(&canonical);
            for lowered_name in [canonical, rooted] {
                self.emit_u16(Op::LOCAL_GET, matched_slot);
                self.emit(Op::I32_EQZ);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_raw_string_slot_eq_literal(callee_name_slot, lowered_name.as_str());
                let line = self.line;
                self.chunk().emit_if(line);

                if let Some(callable_global) =
                    self.source_function_callable_global_name_for_canon(&function_name)
                {
                    self.emit_global_read(&callable_global);
                    self.emit_u16(Op::LOCAL_SET, callee_slot);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, matched_slot);
                }
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
            }
        }
        self.chunk().emit_end(line);
    }
}

/// Emit the call opcode after function ref + args are on stack.
/// Stack before: [func_ref, arg1, arg2, ...]  Stack after: [return_value]
pub fn emit_call(chunk: &mut Chunk, arg_count: u8, line: u32) {
    chunk.emit_op_u8_u8(Op::CALL_REF, arg_count, 1, line);
}

// ── Async/await (WASM Stack Switching + JSPI) ───────────────────────────
//
// All languages use the same async pattern:
//
//   async function:
//     1. Create continuation from body function (cont_new)
//     2. Return the continuation as a Promise-like value
//     3. The runtime schedules it on the event loop
//
//   await expression:
//     1. Compile the expression (produces a value or Promise)
//     2. Emit the spec stack-switching `suspend` instruction tagged with
//        AWAIT_SUSPEND_TAG (JSPI is stack switching applied to JS Promises) —
//        the VM checks if it's a Promise and suspends the fiber if pending.
//        No custom opcode is involved.
//
// Python `async def`, Dart `async`, JS `async function`, C# `async Task`
// all compile to the same opcodes.

/// Module/name of the JSPI suspending import that `await` lowers to. JSPI
/// (JS Promise Integration — stack switching at the JS-promise boundary) marks
/// an import as `WebAssembly.Suspending`; calling it suspends the computation
/// until the returned Promise settles. The VM recognises this import as the
/// suspender (the embedder-side marking) and runs the await/suspend logic.
pub const JSPI_SUSPEND_MODULE: &str = "jspi";
pub const JSPI_SUSPEND_NAME: &str = "await";

/// Emit an `await` expression via JSPI — a plain `call` to a suspending import.
/// Caller must have compiled the awaited expression onto the stack.
/// Stack before: `[value_or_promise]`  Stack after: `[resolved_value]`
pub fn emit_await(chunk: &mut Chunk, line: u32) {
    // Per the JSPI proposal, the suspend point is a normal `call` to a
    // `WebAssembly.Suspending`-marked import — NOT a custom opcode and NOT a
    // magic `suspend` tag. `await x` → `call $jspi.await(x)`, which lowers to
    // the core `call` (0x10) — valid `.wasm`. The VM treats this import as the
    // suspender: fulfilled → unwrap, rejected → throw, pending → suspend the
    // fiber on the event loop (the engine-internal stack switch JSPI mandates)
    // until the Promise settles, then resume with its value. A non-Promise
    // value passes straight through (proposal §"nosuspend").
    let idx = chunk.add_import(JSPI_SUSPEND_MODULE, JSPI_SUSPEND_NAME);
    chunk.emit_call(idx, 1, line); // argc = 1 (the awaited value)
}

/// Two-chunk `await` for runtime-helper builders: the awaited-value `call` is
/// emitted into `code`, but the `jspi.await` import is registered on `imports`
/// (chunks[0]) — matching how those builders register every other import.
/// Adding it to `code`'s own import list instead would shift `code`'s import
/// indices and mis-resolve its other `CALL_IMPORT`s.
pub fn emit_await_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let idx = code.add_import(JSPI_SUSPEND_MODULE, JSPI_SUSPEND_NAME);
    code.emit_call(idx, 1, line); // argc = 1 (the awaited value)
}

/// Emit async function wrapper: wraps the body chunk as a continuation.
/// Call this INSTEAD of the normal function body compilation for async functions.
///
/// The pattern:
///   1. The outer function creates a continuation from the body chunk
///   2. Returns a Promise that resolves when the continuation completes
///
/// `body_chunk_idx`: the chunk index containing the compiled async body
/// Stack: [] → [promise]
pub fn emit_async_wrapper(chunk: &mut Chunk, body_chunk_idx: usize, line: u32) {
    // Create continuation from the body function
    chunk.emit_op_u16(Op::REF_FUNC, body_chunk_idx as u16, line);
    chunk.emit(0, line); // 0 upvalues
    chunk.emit_op(Op::CONT_NEW, line);
    // Resume the continuation immediately — it will suspend at each await point
    // The VM's event loop handles re-resumption when promises resolve
    crate::primitives::generators::emit_resume(chunk, line);
}

/// Create an async function body chunk.
/// Same as create_function_chunk but named with $async suffix for debugging.
pub fn create_async_body_chunk(name: &str, arity: u8) -> Chunk {
    create_function_chunk(&format!("{}$async", name), arity)
}

// ── Spread arguments ───────────────────────────────────────────────────
//
// When a call has spread arguments: f(a, ...arr, b)
// The compiler builds an args array at runtime:
//   1. array_new 0 (empty array)
//   2. For each normal arg: compile + array_push
//   3. For each spread arg: compile + array_concat (flattens into the array)
//   4. Use the array length as argc for the call
//
// This is language-agnostic — JS, Python (*args), Ruby (*splat) all use this.

/// Emit: push one argument onto a spread-args array.
/// Stack before: [args_array, value]  Stack after: [args_array]
///
/// Routes through `ecma:array.push` (returns new length per
/// ECMA-262); caller stashes arr in a local before the loop and
/// reloads afterwards — see `compile_function_decl` rest-args for
/// the canonical template. This helper assumes caller has the stack
/// preserved via a local.
pub fn emit_spread_push_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::primitives::collections::emit_push(chunks, current, line);
}

/// Emit: concat a spread array into the args array via
/// `ecma:array.concat` — returns a new array; caller replaces
/// the accumulator local with the result.
pub fn emit_spread_concat_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::primitives::collections::emit_concat(chunks, current, line);
}
