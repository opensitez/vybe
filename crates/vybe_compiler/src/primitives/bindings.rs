//! Pascal set promotion, var-set, global-map emit, closure/upvalue binding, class-field checks.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use crate::primitives::class_slots;
use super::*;

/// The enclosing frame's closure bookkeeping, held across a nested
/// function-like frame. See [`Compiler::enter_closure_frame`].
pub(super) struct ClosureFrameBooks {
    closure_captured: HashSet<String>,
    env_names: Vec<String>,
    capture_locals: HashMap<u8, u16>,
    shared_env_slot: Option<u16>,
    shared_env_names: Vec<String>,
}

/// Why a bare name is being loaded. The two uses resolve identically except
/// under a closed scope, where the VARIABLE namespace stops chaining outward
/// but the function namespace does not — see [`Compiler::emit_callee_get`].
#[derive(Clone, Copy, PartialEq, Eq)]
enum NameUse {
    Variable,
    Callee,
}

impl Compiler {
    /// An array literal assigned to a `set`-typed binding builds a SET.
    ///
    /// No language check: `hint_is_builtin_set` consults
    /// `profile.builtin_type_spellings`, so only a language that DECLARES a set
    /// spelling can reach the body at all.
    pub(super) fn maybe_promote_array_literal_to_set(
        &mut self,
        type_hint: Option<&str>,
        value: &Expression,
    ) {
        if !type_hint.is_some_and(|h| self.hint_is_builtin_set(h)) {
            return;
        }
        if !matches!(value.kind, ExprKind::Array(_)) {
            return;
        }
        common::sets::emit_from_iterable(&mut self.chunks, self.current, self.line);
    }

    /// Whether `hint` names a `set` ACCORDING TO THE LANGUAGE — its
    /// `[builtin_types] set = [...]` spellings, then the platform table.
    ///
    /// Replaces `Self::is_pascal_set_type_hint`, whose body was the single
    /// spelling `"set of "` living in shared code. Both normalize identically
    /// (`trim().to_lowercase()`) and `set = ["set of *"]` is a `Match::Prefix`,
    /// so for Pascal this is the same predicate — the spelling just moved to
    /// the language that owns it (builtinslotplan.md step 4a).
    pub(super) fn hint_is_builtin_set(&self, hint: &str) -> bool {
        vybe_ast::builtin_types::classify_with(&self.profile.builtin_type_spellings, hint)
            == Some(vybe_ast::builtin_slots::BuiltinType::Set)
    }

    /// Whether `hint` names a plain NUMBER according to the language.
    ///
    /// `BigInt` is deliberately excluded where
    /// `builtin_types::is_numeric` includes it: a bigint is its own runtime
    /// value, not the `Value::Number` that the "reaches only type 0" argument
    /// in `emit_rich_compare_locals` enumerates. Going through
    /// `classify_with` is what makes that exclusion work at all — Kotlin
    /// declares `[builtin_types] bigint = ["long"]`, so its `Long` resolves to
    /// BigInt here and the platform table's `long → Int` row never applies.
    pub(super) fn hint_is_builtin_number(&self, hint: &str) -> bool {
        matches!(
            vybe_ast::builtin_types::classify_with(&self.profile.builtin_type_spellings, hint),
            Some(
                vybe_ast::builtin_slots::BuiltinType::Int
                    | vybe_ast::builtin_slots::BuiltinType::Double
            )
        )
    }

    /// Whether `hint` names a `string` ACCORDING TO THE LANGUAGE.
    ///
    /// Distinct from `Self::is_string_type_hint`, which consults the PLATFORM
    /// table only. Pascal's `char` holds a character — its default value is
    /// `""` and `Ord`/`Chr` round-trip through it — where the C family's `char`
    /// is an 8-bit integer. Pascal says so with `[builtin_types] string =
    /// ["char"]` rather than the shared compiler naming it (§3c).
    pub(super) fn hint_is_builtin_string(&self, hint: &str) -> bool {
        vybe_ast::builtin_types::classify_with(&self.profile.builtin_type_spellings, hint)
            == Some(vybe_ast::builtin_slots::BuiltinType::String)
    }

    /// Whether `expr` is statically known to be a set.
    ///
    /// Carries NO language name. It used to open with `profile.name !=
    /// "pascal"`, which made it a Pascal predicate that happened to live in
    /// shared code; reachability is now the CALLER's business, and every caller
    /// is gated on the language having declared a set binding (§2d — a language
    /// that declares nothing cannot reach these paths).
    pub(super) fn expr_is_builtin_set(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Set(_) => true,
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .is_some_and(|h| self.hint_is_builtin_set(&h)),
            ExprKind::Binary { op, left, right }
                if matches!(op, BinOp::Add | BinOp::Mul | BinOp::Sub) =>
            {
                self.expr_is_builtin_set(left) && self.expr_is_builtin_set(right)
            }
            _ => false,
        }
    }

    pub(crate) fn emit_var_get(&mut self, name: &str) {
        self.emit_name_get(name, NameUse::Variable);
    }

    /// Load `name` as the TARGET OF A CALL rather than as a variable read.
    ///
    /// The two differ in exactly one place — the closed-scope rule below. A
    /// closed scope (PHP's) nulls a name that is not one of its own locals,
    /// because the VARIABLE namespace does not chain outward; that rule is what
    /// makes `$nope` inside a function read NULL even when a top-level `$nope`
    /// exists, and it must not change. A callee is not a variable read: the
    /// rule already tries to exempt functions, which live in one flat namespace,
    /// but it tests `defined_functions` — a COMPILE-TIME snapshot. A function
    /// published by a runtime `include`/`require` (or `eval`) is not in that
    /// snapshot, so its callee was nulled before it could ever be found.
    pub(crate) fn emit_callee_get(&mut self, name: &str) {
        self.emit_name_get(name, NameUse::Callee);
    }

    fn emit_name_get(&mut self, name: &str, use_site: NameUse) {
        // Shared env: locals captured by inner closures live in a shared
        // array so mutations are visible across all closures.
        if let Some(idx) = self.shared_env_index(name) {
            if let Some(env_slot) = self.shared_env_slot {
                let l = self.line;
                crate::primitives::closures::emit_env_get(self.chunk(), env_slot, idx, l);
                return;
            }
        }
        // Local
        if let Some(slot) = self.scope().resolve(name) {
            self.emit_u16(Op::LOCAL_GET, slot);
            if self.binding_uses_pointer_cell(name) {
                // Not `emit_cell_load` — a binding can hold EITHER reference
                // shape (`&$x` gives a cell, `&$a[1]` gives a carray) and the
                // cell-only load read `__value` off a carray, i.e. undefined.
                self.emit_autoderef_pointer_cell();
            }
            return;
        }
        if self.scopes.len() > 1 {
            if let Some(_uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index(name);
                let l = self.line;
                crate::primitives::closures::emit_env_get(self.chunk(), env, idx, l);
                return;
            }
        }
        if let Some(binding) = self.static_local_binding(name) {
            let global_name = binding.global_name.clone();
            self.emit_global_read(&global_name);
            return;
        }
        // Implicit self field — when inside a class method and the name is a
        // field of the current class, read from `me.<name>`. This is what
        // languages like VB do for unqualified field access. Without this,
        // dotted-name resolution that returns InstanceMember { local: "field" }
        // would fall through to global_get and read null.
        if self.current_class_implicit_self && self.is_class_field(name) {
            if self.emit_self_ref() {
                let cname = self
                    .current_class
                    .as_deref()
                    .and_then(|class_name| {
                        self.visible_instance_field_storage_name_for_class(class_name, name)
                    })
                    .unwrap_or_else(|| self.canon(name));
                self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal(&cname));
                return;
            }
        }
        // Static field of the current class — `Count++` inside `Counter`
        // ctor reads `Counter.Count` (struct_get on the class global).
        // Without this, the bare name falls through to global_get and
        // returns null because the static field lives on the class
        // struct, not the module's global namespace.
        if let Some(class_name) = self.is_class_static_field(name) {
            self.emit_global_read(&class_name);
            self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal(&self.canon(name)));
            return;
        }
        // Bare static method in class scope — `Double(x)` inside
        // `class Converter` resolves to `Converter.Double`.
        if let Some(class_name) = self.is_class_static_method(name) {
            self.emit_global_read(&class_name);
            self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal(&self.canon(name)));
            return;
        }
        let cname = self.canon(name);
        let shadows_named_global = self.defined_globals.contains(&cname)
            || self.defined_functions.contains(&cname)
            || self.defined_classes.contains(&cname);
        if !shadows_named_global && self.emit_with_target_get(name) {
            return;
        }
        // Known type used as a value (e.g. `e instanceof RangeError`) — emit
        // the type name as a string so the instanceof ref.test fallback can
        // look it up by name. Without this, `RangeError` would become
        // `global_get` of a nonexistent global → null.
        // Only do this when the name isn't shadowed by an actual global
        // (e.g. `Dim list As New List(Of String)` shadows the `list` type name).
        let is_js_runtime_global = self.profile.has_ecma_globals
            && (matches!(
                name,
                "Object"
                    | "Boolean"
                    | "Number"
                    | "String"
                    | "Array"
                    | "Function"
                    | "Symbol"
                    | "BigInt"
                    | "Error"
                    | "EvalError"
                    | "RangeError"
                    | "ReferenceError"
                    | "SyntaxError"
                    | "TypeError"
                    | "URIError"
                    | "AggregateError"
                    | "SuppressedError"
                    | "ArrayBuffer"
                    | "SharedArrayBuffer"
                    | "DataView"
                    | "Int8Array"
                    | "Uint8Array"
                    | "Uint8ClampedArray"
                    | "Int16Array"
                    | "Uint16Array"
                    | "Int32Array"
                    | "Uint32Array"
                    | "Float32Array"
                    | "Float64Array"
                    | "BigInt64Array"
                    | "BigUint64Array"
            ) || self.host_namespace_aliases.contains_key(&cname));
        if self.profile.known_types.contains_key(name)
            && !self.defined_globals.contains(name)
            && !self.defined_globals.contains(&cname)
            && !is_js_runtime_global
        {
            self.emit_const(Value::String(Arc::from(name)));
            return;
        }
        // JS builtin constructor used as a *value* (`x === Array`,
        // `o.constructor === Object`, `Array.prototype`): resolve through the
        // stable, host-owned `__ctor_<Name>` anchor instead of the user-facing
        // global. The latter can be re-bound to a fresh, unwired object by
        // later compile/link passes (ESM import wiring), which would break
        // `constructor`/`prototype` identity; `__ctor_<Name>` always points at
        // the ONE canonical constructor (the same object on the shared
        // prototype's `.constructor`). Skipped when the user shadows the name.
        // Built-in Error constructors are recognised from the profile's
        // `known_types` (their backing module is `ecma:error`) rather than a
        // hardcoded name list, so `e.constructor === TypeError` and
        // `typeof TypeError === "function"` resolve through the same canonical
        // `__ctor_<Name>` anchor the host installs for them.
        let is_error_ctor_value = self
            .profile
            .known_types
            .get(name)
            .is_some_and(|(module, _)| module == "ecma:error");
        if self.profile.has_ecma_globals
            && !shadows_named_global
            && (is_js_builtin_ctor_value(&cname) || is_error_ctor_value)
        {
            self.emit_global_read(&format!("__ctor_{cname}"));
            return;
        }
        // A CLOSED scope does not chain outward: the name is not a local, so it
        // reads null rather than resolving to a module global. Functions and
        // classes stay reachable (they live in one flat namespace, not the
        // variable one), as do compiler internals and use-aliases.
        //
        // `use const Lib\LEVEL;` imported names read the qualified global from
        // inside a closed scope too — fall through to the use-alias consult
        // below instead of the undeclared-null.
        //
        // A CALLEE falls through to the global read instead: reading the name
        // is late-bound, so an undefined function is still null at runtime
        // (unchanged), while one defined a moment ago by a runtime include
        // now resolves. See [`Self::emit_callee_get`].
        if use_site == NameUse::Variable
            && !self.scope().is_open(&cname)
            && !self.defined_functions.contains(&cname)
            && !self.defined_classes.contains(&cname)
            && !cname.starts_with("__")
            && !self.source_type_aliases.contains_key(&cname)
        {
            self.emit_null();
            return;
        }
        // Global — canonicalize name for case-insensitive languages
        // But in strict mode, if this is genuinely undeclared, throw ReferenceError
        //
        // Use-alias consult (namespaceplan.md): `use const Lib\MAX;` binds
        // bare `MAX` to the namespace-qualified global `Lib.MAX` (the same
        // `source_type_aliases` map static-access/instanceof already
        // resolve through). Only when nothing else declared the bare name —
        // a real global/function/class shadows the import per PHP scoping.
        let cname = if self.profile.uses_common_resolver
            && !self.defined_globals.contains(&cname)
            && !self.defined_functions.contains(&cname)
            && !self.defined_classes.contains(&cname)
        {
            // PHP §namespace resolution order for unqualified names: the
            // CURRENT namespace's declaration wins over file-level `use`
            // aliases (which are last-wins across the file), then the
            // aliases, then the global fallback.
            let ns_qualified = self.current_namespace.as_deref().and_then(|ns| {
                let q = self.canon(&format!("{ns}.{cname}"));
                (self.defined_functions.contains(&q) || self.defined_globals.contains(&q))
                    .then_some(q)
            });
            match ns_qualified {
                Some(q) => q,
                None => match self.source_type_aliases.get(&cname) {
                    Some(target) => self.canon(target),
                    None => self.resolve_source_namespace_value(&cname).unwrap_or(cname),
                },
            }
        } else {
            cname
        };
        let global_key = self.variable_global_key(name, &cname);

        // ECMA-262 §9.1.1.4.6 / §13.3.2.1 GetValue: reading an *unresolvable*
        // reference (a name bound nowhere in the scope chain or on the global
        // object) is a `ReferenceError`. Reaching this fallback means every
        // compile-time resolution attempt failed — not a local/upvalue/static/
        // class-field, not a declared global/function/class, not a builtin or
        // host target (those returned earlier). So this is a genuine *missing
        // binding*, decided at compile time — NOT a runtime "value is undefined"
        // test (a declared `let x;` legitimately holds `undefined`).
        //
        // Driven by the `unresolved_reference_throws` profile capability, and
        // additionally gated on strict mode: per spec the throw applies in
        // sloppy mode too, but sloppy code leans on lenient access to
        // host-provided globals the compiler does not track as bindings (so a
        // blanket throw there mis-fires). `typeof x` on an undeclared name must
        // yield "undefined", so it is suppressed via `in_typeof_operand`.
        // Builtin global-object aliases are excluded defensively.
        let resolvable = self.defined_globals.contains(&cname)
            || self.defined_functions.contains(&cname)
            || self.defined_classes.contains(&cname);
        // §9.1.1.4.6 GetValue throws for an unresolvable read in BOTH strict
        // and sloppy mode. We gate the general throw to strict mode because a
        // blanket sloppy throw misfires on host globals we don't track as
        // bindings — but a name the program itself lexically declared
        // (`program_lexical_names`) that is unresolvable here is provably an
        // out-of-scope *user* binding, never an untracked host global, so it
        // throws in sloppy mode too.
        // MEASURED: dropping the strict/lexical-name half for a non-ECMA
        // profile cost 47 python tests (basics/classes/control_flow/
        // comprehensions/closure_extended went 130/7 → 83/54). The carve-out is
        // not an ECMA quirk — plenty of names resolve at runtime that the
        // compiler never saw bound, in every language. Left as it was.
        if self.profile.unresolved_reference_throws
            && (self.in_strict || self.program_lexical_names.contains(&cname))
            && !self.in_typeof_operand
            && !resolvable
            && !cname.starts_with("__")
            && !is_js_builtin_ctor_value(&cname)
            && !matches!(
                name,
                "globalThis"
                    | "window"
                    | "self"
                    | "global"
                    | "globalObject"
                    | "arguments"
                    | "this"
                    | "undefined"
                    | "NaN"
                    | "Infinity"
            )
        {
            let line = self.line;
            self.class_alloc();
            inst!(self, core_wasm::dup);
            // Exception type and message text come from the profile — JS
            // `ReferenceError: x is not defined`, Python
            // `NameError: name 'x' is not defined`. The structured-exception
            // machinery in `primitives/errors.rs` already takes the kind as a
            // parameter; only these two literals were hardcoded.
            let message = self
                .profile
                .unresolved_reference_message
                .replace("{}", name);
            let error_kind = self.profile.unresolved_reference_error.clone();
            self.emit_const(Value::String(Arc::from(message.as_str())));
            crate::primitives::errors::emit_exception_new_finalize(self.chunk(), &error_kind, line);
            crate::primitives::errors::emit_throw(self.chunk(), line);
            return;
        }
        self.emit_global_read(&global_key);
        if self.binding_uses_pointer_cell(name) {
            // Either reference shape — see the local arm above.
            self.emit_autoderef_pointer_cell();
        }
    }

    /// Same as `emit_ensure_global_map` but for an ARRAY-valued global.
    pub(super) fn emit_ensure_global_list(&mut self, name: &str) {
        self.emit_global_read(name);
        inst!(self, core_wasm::dup);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit(Op::DROP);
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        inst!(self, core_wasm::dup);
        self.emit_global_write(name);

        self.chunk().emit_end(line);
        self.emit(Op::DROP);
    }

    /// Run every `register_shutdown_function` callback, in registration order,
    /// then clear the list so a later `exit` cannot run them twice.
    ///
    /// Emitted at the normal end of a PHP program AND immediately before
    /// `exit`/`die` terminate it — real php runs shutdown handlers on both
    /// paths, which is what makes a check registered this way survive an
    /// `exit(1)` in the middle of a script.
    pub(super) fn emit_php_run_shutdown_fns(&mut self) {
        let line = self.line;

        // `REF_IS_NULL` leaves a RAW i32; running the boxed `dyn_not` /
        // `dyn_to_bool` on it corrupts the stack (measured: the handlers ran,
        // then the program threw `RuntimeError: [object]`). Branch on it
        // directly and put the work in the ELSE arm, the way
        // `emit_ensure_global_map` does.
        self.emit_global_read("__php_shutdown_fns");
        self.emit(Op::REF_IS_NULL);
        self.chunk().emit_if(line);
        self.chunk().emit_else(line);

        // `alloc_scratch`, NOT `define_local`: this helper runs from the
        // module EPILOGUE, after the scope's slot count is settled, so a
        // scope-allocated local is never reserved in the call frame and reads
        // back garbage. Raw chunk slots are folded into the frame size by the
        // `max` the epilogue takes right after.
        let list = self.chunk().alloc_scratch(1);
        self.emit_global_read("__php_shutdown_fns");
        self.emit_u16(Op::LOCAL_SET, list);
        // Clear FIRST: a handler that itself calls `exit` would otherwise
        // re-enter this and run the whole list again.
        self.emit_null();
        self.emit_global_write("__php_shutdown_fns");

        let idx = self.chunk().alloc_scratch(1);
        self.emit_const(Value::F64(0.0));
        self.emit_u16(Op::LOCAL_SET, idx);

        let block = self.chunk().emit_block(line);
        let (loop_patch, _) = self.chunk().emit_loop_s(line);

        self.emit_u16(Op::LOCAL_GET, idx);
        self.emit_u16(Op::LOCAL_GET, list);
        common::collections::emit_len(&mut self.chunks, self.current, line);
        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        self.chunk().emit_br_if(1, line);

        // entry = list[i]; call entry[0] with entry[1..]
        let entry = self.chunk().alloc_scratch(1);
        self.emit_u16(Op::LOCAL_GET, list);
        self.emit_u16(Op::LOCAL_GET, idx);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        self.emit_u16(Op::LOCAL_SET, entry);

        // `apply(fn, thisArg, argsArray)` — THREE arguments, as every other
        // call site emits. Passing two consumed the argument array as
        // `thisArg`, so a handler registered with extra arguments
        // (`register_shutdown_function($fn, "ARG")`) was called with none:
        // `SHUTDOWN ` where php prints `SHUTDOWN ARG`.
        self.emit_u16(Op::LOCAL_GET, entry);
        self.emit_const(Value::F64(0.0));
        common::collections::emit_get(&mut self.chunks, self.current, line);
        self.emit_null();
        self.emit_u16(Op::LOCAL_GET, entry);
        self.emit_const(Value::F64(1.0));
        let slice = self.import("ecma:array", "slice");
        self.emit_host_call(slice, 2);
        let apply = self.import("ecma:function", "apply");
        self.emit_host_call(apply, 3);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, idx);
        self.emit_const(Value::F64(1.0));
        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, idx);

        self.chunk().emit_br(0, line);

        // THREE structural regions are open here — the `if`, the `block` and
        // the `loop` — so three `end`s close them. `patch_loop`/`patch_block`
        // are no-ops (the block table replaced size-header patching), so they
        // do NOT close anything; emitting one `end` for three regions left the
        // `if` unterminated, and its else-arm then ran unconditionally. That
        // made every program in every language execute this php-only runner
        // over an absent `__php_shutdown_fns`, ending in
        // `ecma:object.keys(undefined)` — `RuntimeError: [object]` after
        // otherwise-correct output.
        self.chunk().emit_end(line); // close loop
        self.chunk().patch_loop(loop_patch);
        self.chunk().emit_end(line); // close block
        self.chunk().patch_block(block);
        self.chunk().emit_end(line); // close if/else
    }

    pub(super) fn emit_ensure_global_map(&mut self, name: &str) {
        self.emit_global_read(name);
        inst!(self, core_wasm::dup);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit(Op::DROP);
        common::collections::emit_map_new(&mut self.chunks, self.current, line);
        inst!(self, core_wasm::dup);
        self.emit_global_write(name);

        self.chunk().emit_end(line);
    }
    /// ECMA-262 §11.2.1 Directive Prologue: returns `true` if the leading
    /// run of string-literal expression statements contains `"use strict"`.
    pub(crate) fn stmts_have_use_strict_directive(stmts: &[Statement]) -> bool {
        for s in stmts {
            match &s.kind {
                // The walker emits `Empty` for newlines between statements;
                // they don't terminate the directive prologue.
                StmtKind::Empty => continue,
                // The JS walker HOISTS function declarations above the
                // directive prologue (parse-time reorder), so a decl here
                // says nothing about the prologue's textual position.
                StmtKind::FunctionDecl { .. } => continue,
                StmtKind::Expr(e) => match &e.kind {
                    ExprKind::Lit(Literal::Str(v)) => {
                        if v == "use strict" {
                            return true;
                        }
                        // Another directive — keep scanning the prologue.
                    }
                    _ => break,
                },
                _ => break,
            }
        }
        false
    }

    /// Store the value on the stack into `name`, writing THROUGH the reference
    /// if the binding holds one.
    pub(super) fn emit_var_set(&mut self, name: &str) {
        self.emit_var_store(name, true);
    }

    /// BIND `name` to the value on the stack — the value IS what the name now
    /// denotes, even when that value is a reference.
    ///
    /// The difference from [`Self::emit_var_set`] is the whole of php's
    /// `$b = &$a`. Storing THROUGH would ask `binding_uses_pointer_cell`, and
    /// when the module-wide address-taken pre-pass says a wrap is still coming
    /// for `$b` that answers `true` before any cell exists — so the store mints
    /// a fresh cell and puts `$a`'s reference INSIDE it, orphaning the storage
    /// the reference names. `$b` and `$c` then share the outer cell while `$a`
    /// keeps the old value: the reference-chain defect.
    ///
    /// The mark comes AFTER the store, and the order is load-bearing: a first
    /// assignment is what CREATES the binding, so marking ahead of it finds no
    /// local to flag and records the fact on the module-wide global store
    /// instead. The local is then born unmarked, its reads never deref, and the
    /// interpolation gets a cell object where it wanted a number. The store
    /// itself never consults the mark, so nothing needs it earlier.
    pub(super) fn emit_var_bind_reference(&mut self, name: &str) {
        self.emit_var_store(name, false);
        self.mark_pointer_cell_binding(name);
    }

    fn emit_var_store(&mut self, name: &str, through_reference: bool) {
        // ECMA-262 §13.15.2 / §6.2.4.7: assigning to a `const` binding is a
        // runtime `TypeError` ("Assignment to constant variable."). The
        // binding is known to the compiler — a `const` local in scope, or a
        // top-level `const` global — so emit an unconditional throw at the
        // assignment site. (Declaration init and direct loop-variable rebinds
        // use `LOCAL_SET`/`GLOBAL_SET` directly and never reach here.)
        if self.profile.ecma_lexical_declarations {
            let is_const_local = self.scope().resolve_is_const(name);
            let is_const_global = !is_const_local
                && self.scope().resolve(name).is_none()
                && (self.const_globals.contains(name)
                    || self.const_globals.contains(&self.canon(name)));
            if is_const_local || is_const_global {
                let line = self.line;
                self.class_alloc();
                inst!(self, core_wasm::dup);
                self.emit_const(Value::String(Arc::from("Assignment to constant variable.")));
                crate::primitives::errors::emit_exception_new_finalize(
                    self.chunk(),
                    "TypeError",
                    line,
                );
                crate::primitives::errors::emit_throw(self.chunk(), line);
                return;
            }
        }
        // Shared env: locals captured by inner closures
        if let Some(idx) = self.shared_env_index(name) {
            if let Some(env_slot) = self.shared_env_slot {
                let l = self.line;
                crate::primitives::closures::emit_env_set(self.chunk(), env_slot, idx, l);
                return;
            }
        }
        // Local
        if let Some(slot) = self.scope().resolve(name) {
            if through_reference && self.binding_uses_pointer_cell(name) {
                let value_slot = self.define_local("__ref_cell_set_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                // Same rule as the global arm below: the pre-pass may be
                // answering for a wrap that is still ahead, and the first write
                // must CREATE the cell rather than store through a missing one.
                if !self.binding_already_pointer_cell(name) {
                    self.promote_local_binding_to_pointer_cell(name);
                }
                // Shape-polymorphic, matching the load above: a cell-only store
                // wrote `__value` onto a carray, growing a dead field while the
                // container it referenced was never touched.
                self.emit_store_through_pointer(slot, value_slot);
            } else if let Some((args_slot, index)) = self.js_arguments_alias_for_name(name) {
                let value_slot = self.define_local("__js_arguments_alias_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_u16(Op::LOCAL_SET, slot);
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.emit_const(Value::F64(index as f64));
                self.emit_u16(Op::LOCAL_GET, value_slot);
                common::collections::emit_set(&mut self.chunks, self.current, self.line);
                self.emit(Op::DROP);
            } else {
                self.emit_u16(Op::LOCAL_SET, slot);
            }
            return;
        }
        if self.scopes.len() > 1 {
            if let Some(_uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index(name);
                let l = self.line;
                crate::primitives::closures::emit_env_set(self.chunk(), env, idx, l);
                return;
            }
        }
        if let Some(binding) = self.static_local_binding(name) {
            let global_name = binding.global_name.clone();
            self.emit_global_write(&global_name);
            return;
        }
        if self.current_class_implicit_self && self.is_class_field(name) {
            let value_slot = self.define_local("__implicit_self_value");
            self.emit_u16(Op::LOCAL_SET, value_slot);
            if self.emit_self_ref() {
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let cname = self
                    .current_class
                    .as_deref()
                    .and_then(|class_name| {
                        self.visible_instance_field_storage_name_for_class(class_name, name)
                    })
                    .unwrap_or_else(|| self.canon(name));
                self.class_set(
                    class_slots::ObjSource::Stack,
                    &class_slots::ClassSlot::internal(&cname),
                    class_slots::ValueSource::Stack,
                );
                return;
            }
            self.emit_u16(Op::LOCAL_GET, value_slot);
        }
        // Static field of the current class — write through to
        // `<ClassName>.<name>` instead of falling to global_set.
        if let Some(class_name) = self.is_class_static_field(name) {
            // Stack: [value]. Need [class_obj, value] for STRUCT_SET.
            let value_slot = self.define_local("__static_set_value");
            self.emit_u16(Op::LOCAL_SET, value_slot);
            self.emit_global_read(&class_name);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            let bare_name = self.canon(name);
            self.class_set(
                class_slots::ObjSource::Stack,
                &class_slots::ClassSlot::internal(&bare_name),
                class_slots::ValueSource::Stack,
            );
            if self.defined_globals.contains(&bare_name) {
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_global_write(&bare_name);
            }
            return;
        }
        let cname = self.canon(name);
        let shadows_named_global = self.defined_globals.contains(&cname)
            || self.defined_functions.contains(&cname)
            || self.defined_classes.contains(&cname);
        // ECMA-262 §6.2.5.6 PutValue / §9.1.1.4.5: in strict mode, assigning
        // to an unresolvable reference (a name bound nowhere in the scope
        // chain or on the global object) is a `ReferenceError` — sloppy mode
        // would silently create a global. Reaching here means the name is not
        // a local/upvalue/static/class-field/declared-global. We can only
        // throw for names that cannot be a host-provided builtin global, so
        // exclude the known builtin constructors and global-object aliases.
        // Gated on `in_strict` (rare) to keep sloppy-mode global-creation,
        // which the rest of the suite relies on, intact.
        if self.in_strict
            && self.profile.ecma_strict_mode
            && !shadows_named_global
            && !cname.starts_with("__")
            && !is_js_builtin_ctor_value(&cname)
            && !matches!(
                name,
                "globalThis" | "window" | "self" | "global" | "globalObject" | "arguments"
            )
        {
            let line = self.line;
            self.class_alloc();
            inst!(self, core_wasm::dup);
            self.emit_const(Value::String(Arc::from(
                format!("{name} is not defined").as_str(),
            )));
            crate::primitives::errors::emit_exception_new_finalize(
                self.chunk(),
                "ReferenceError",
                line,
            );
            crate::primitives::errors::emit_throw(self.chunk(), line);
            return;
        }
        if !shadows_named_global && self.emit_with_target_set(name) {
            return;
        }
        // Closed scope: an assignment to a name that isn't open CREATES a local
        // here rather than writing the module global.
        if !self.scope().is_open(&cname)
            && !self.defined_functions.contains(&cname)
            && !self.defined_classes.contains(&cname)
            && !cname.starts_with("__")
        {
            let slot = self.define_source_local(name);
            self.emit_u16(Op::LOCAL_SET, slot);
            return;
        }
        // Global — canonicalize name for case-insensitive languages
        let global_key = self.variable_global_key(name, &cname);
        if self.scopes.len() == 1 {
            self.defined_globals.insert(global_key.clone());
        }
        if through_reference && self.binding_uses_pointer_cell(name) {
            let value_slot = self.define_local("__ref_global_set_value");
            self.emit_u16(Op::LOCAL_SET, value_slot);
            // The module-wide pre-pass answers `true` before any wrap has
            // happened, and at module scope the FIRST write is what gives the
            // global its value — `$a = 1;` ahead of the `inc($a)` that promotes
            // it. Storing through a cell that does not exist yet drops the
            // value silently, and every later read of the alias sees undefined.
            //
            // Creating it here IS the "promote once, at declaration time" the
            // pre-pass exists to schedule; the promotion marks itself, so the
            // wrap still happens exactly once. Emitted with the value already
            // off the stack, because the promotion emits code of its own.
            if !self.binding_already_pointer_cell(name) {
                self.promote_global_binding_to_pointer_cell(name);
            }
            let ptr_slot = self.define_local("__ref_global_set_ptr");
            self.emit_global_read(&global_key);
            self.emit_u16(Op::LOCAL_SET, ptr_slot);
            // Either reference shape — see the local arm above.
            self.emit_store_through_pointer(ptr_slot, value_slot);
            return;
        }
        self.emit_global_write(&global_key);
    }

    /// Enter a function-like FRAME's closure bookkeeping, returning the
    /// enclosing frame's books for [`Self::exit_closure_frame`].
    ///
    /// Five fields are per-frame: the captured-locals scan, the closure env
    /// NAME LAYOUT, the upvalue slot map, and the shared-env slot/name pair.
    /// Every site that pushes a function chunk + scope must reset all five —
    /// this pair is the ONE home for that ritual. It used to be spelled
    /// inline at each site, and the copies drifted exactly as duplicated
    /// mechanisms do: the anonymous-class site forgot `capture_locals`, and
    /// the by-value capture factory reset nothing at all, so an inner
    /// lambda's `closure_env_index` kept answering from the ENCLOSING
    /// function's name list — `env[N]` of a smaller env, null captures in
    /// every language that declares bare-name captures.
    ///
    /// `seed_env_names`: a frame whose parent carries a shared env pre-seeds
    /// its own name layout with the parent's, so upvalue indices computed
    /// inside line up with the parent's shared-env array. Pass the parent's
    /// `shared_env_names` (or `&[]` when the frame has no such parent — the
    /// capture factory, whose only locals are the by-value captures).
    pub(super) fn enter_closure_frame(&mut self, seed_env_names: &[String]) -> ClosureFrameBooks {
        let books = ClosureFrameBooks {
            closure_captured: std::mem::take(&mut self.current_closure_captured_locals),
            env_names: std::mem::take(&mut self.closure_env_names),
            capture_locals: std::mem::take(&mut self.capture_locals),
            shared_env_slot: self.shared_env_slot.take(),
            shared_env_names: std::mem::take(&mut self.shared_env_names),
        };
        if !seed_env_names.is_empty() {
            self.closure_env_names = seed_env_names.to_vec();
        }
        books
    }

    /// Restore the enclosing frame's closure books. Call with the value the
    /// matching [`Self::enter_closure_frame`] returned, after `self.current`
    /// is back on the enclosing chunk and BEFORE any code is emitted there —
    /// upvalue env construction in the enclosing frame consults these.
    pub(super) fn exit_closure_frame(&mut self, books: ClosureFrameBooks) {
        self.current_closure_captured_locals = books.closure_captured;
        self.closure_env_names = books.env_names;
        self.capture_locals = books.capture_locals;
        self.shared_env_slot = books.shared_env_slot;
        self.shared_env_names = books.shared_env_names;
    }

    pub(super) fn capture_local_slot(&mut self, uv_idx: u8) -> u16 {
        if let Some(&slot) = self.capture_locals.get(&uv_idx) {
            return slot;
        }
        let slot = self.define_local(&format!("__capture_{}", uv_idx));
        self.capture_locals.insert(uv_idx, slot);
        let c = &mut self.chunks[self.current];
        if c.capture_count <= uv_idx {
            c.capture_count = uv_idx + 1;
        }
        if c.capture_base == 0 || slot < c.capture_base {
            c.capture_base = slot;
        }
        slot
    }

    /// Get or allocate the closure environment slot for the current function.
    /// The env is a GC array holding all captured variables by index.
    /// It arrives as upvalue[0] and is copied to this local by call_function_inner.
    pub(super) fn closure_env_slot(&mut self) -> u16 {
        self.capture_local_slot(0)
    }

    /// Get or register a captured variable's index in the closure env array.
    pub(super) fn closure_env_index(&mut self, name: &str) -> u16 {
        if let Some(idx) = self.closure_env_names.iter().position(|n| n == name) {
            return idx as u16;
        }
        let idx = self.closure_env_names.len();
        self.closure_env_names.push(name.to_string());
        idx as u16
    }

    /// Check if a name is in the current function's shared env.
    pub(super) fn shared_env_index(&self, name: &str) -> Option<u16> {
        self.shared_env_names
            .iter()
            .position(|n| n == name)
            .map(|i| i as u16)
    }

    pub(super) fn resolve_upvalue(&mut self, scope_idx: usize, name: &str) -> Option<u8> {
        if scope_idx == 0 {
            return None;
        }
        let parent = scope_idx - 1;
        // Check parent's locals
        let found_local = self.scopes[parent].resolve(name);
        if let Some(slot) = found_local {
            self.scopes[parent].mark_captured(slot);
            return Some(self.scopes[scope_idx].add_upvalue(slot, true));
        }
        // Recurse up
        if let Some(uv) = self.resolve_upvalue(parent, name) {
            return Some(self.scopes[scope_idx].add_upvalue(uv as u16, false));
        }
        None
    }

    /// Returns the owning class name when `name` is a static field of
    /// the currently-compiling class (or one of its ancestors). Used by
    /// `emit_var_get` / `emit_var_set` to rewrite bare references to
    /// `ClassName.name` so static state lives on the class struct.
    pub(super) fn is_class_static_field(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        if pc.static_fields.iter().any(|f| {
                            if self.case_sensitive {
                                f == name
                            } else {
                                f.eq_ignore_ascii_case(name)
                            }
                        }) {
                            return Some(cn.to_string());
                        }
                        current = pc.parent.as_deref();
                    } else {
                        break;
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }
}
