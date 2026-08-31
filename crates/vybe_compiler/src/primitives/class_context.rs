//! Class-scope membership checks, self/this/new-target handling, member-chain helpers.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use crate::primitives::class_slots;
use super::*;

pub(crate) use vybe_runtime::chunk::ReceiverAbi;

/// The module's receiver ABI, read off the module chunk (`chunks[0]`).
///
/// For the emitters that hold `&mut [Chunk]` and no `Compiler` — see the field
/// docs on [`vybe_runtime::chunk::Chunk::module_receiver_abi`].
pub fn module_receiver_abi(chunks: &[Chunk]) -> ReceiverAbi {
    chunks
        .first()
        .map_or(ReceiverAbi::Ambient, |m| m.module_receiver_abi)
}

/// Bind `recv_slot` as the receiver of the invoke that immediately follows.
///
/// ⛔ Emits the `local.get` TOO, not just the store — under
/// [`ReceiverAbi::Parameter`] both have to disappear together, and a caller
/// that pushed the value itself would leave it stranded on the stack.
///
/// Under `Parameter` this emits NOTHING, and that is the whole point: every
/// site that calls this already pushes the receiver as argument 0 of the
/// invoke (host builtins have always read their receiver there — `ecma:array.map`
/// opens `array_of(args, 0)`). The ambient store was the SECOND copy, kept
/// only for the bytecode path. Deleting it leaves the argument as the single
/// channel, which is ECMA-262 §10.2.1.
/// Takes `abi` rather than `&[Chunk]` because every caller has already
/// destructured `let chunk = &mut chunks[current]` and holds that borrow.
pub fn bind_ambient_receiver(
    chunk: &mut Chunk,
    abi: ReceiverAbi,
    recv_slot: u16,
    line: u32,
) {
    if abi != ReceiverAbi::Ambient {
        return;
    }
    chunk.emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    crate::primitives::globals::emit_write(chunk, "__js_this", line);
}

/// The local frame of a HAND-BUILT stdlib chunk (`build_generator_next` and
/// friends), which have no `Compiler` and no scope to allocate against.
///
/// ⛔ THIS EXISTS BECAUSE THOSE BUILDERS WROTE THEIR LOCAL INDICES AS
/// LITERALS — `let value_local = 0u16; let has_more_local = 1u16;`. Under
/// [`ReceiverAbi::Parameter`] the receiver takes slot 0 and every one of those
/// literals is off by one. Renumbering them by hand means being right at every
/// site in the builder, silently reading a neighbouring local everywhere else;
/// `alloc_scratch` aliasing named locals has already cost one dart test that
/// way. Allocating the slots makes the shift happen ONCE, here.
pub(crate) struct StdlibFrame {
    abi: ReceiverAbi,
    /// Does THIS chunk declare a receiver parameter?
    ///
    /// ⛔ NOT the same question as `abi`. The module's ABI decides how this
    /// chunk CALLS other functions; how this chunk is REACHED decides whether
    /// it declares a parameter of its own, and the two can disagree.
    /// `__stdlib_iter_drain` is the case in point: it invokes user methods, so
    /// it must pass receivers the new way, but its only caller is a hand-built
    /// site in this crate that pushes none — give it a receiver parameter and
    /// its one real call site desyncs.
    declares_receiver: bool,
}

impl StdlibFrame {
    /// For a chunk reached as a METHOD — installed as an object property and
    /// called `o.m()`, so under `Parameter` its call sites push a receiver and
    /// it must declare one. Reserves slot 0 for it.
    ///
    /// Allocate the user parameters first, in declaration order, then the
    /// scratch locals — binding is POSITIONAL, so declaration order is the ABI.
    pub(crate) fn method(abi: ReceiverAbi, chunk: &mut Chunk) -> Self {
        let declares_receiver = abi == ReceiverAbi::Parameter;
        if declares_receiver {
            chunk.alloc_scratch(1); // slot 0, the receiver
        }
        Self {
            abi,
            declares_receiver,
        }
    }

    /// For a chunk reached ONLY from hand-built call sites in this crate,
    /// which push no receiver argument under either ABI. It declares none —
    /// but still CALLS other functions with the module's ABI.
    pub(crate) fn plain(abi: ReceiverAbi) -> Self {
        Self {
            abi,
            declares_receiver: false,
        }
    }

    /// The next slot. Same call for parameters and locals: a WASM frame does
    /// not distinguish them, only the arity split does.
    ///
    /// ⛔ DELEGATES TO THE CHUNK'S OWN ALLOCATOR — it does NOT keep a counter
    /// of its own. A parallel counter is what made the first version of this
    /// wrong: the frame's slots were only published to `chunk.local_count` at
    /// `finish()`, so an emitter helper calling `alloc_scratch` in between
    /// (`emit_next` does, on its very first line) was handed slot 0 and
    /// silently aliased the receiver. With one allocator that cannot happen,
    /// and interleaving a helper's allocations with the frame's is safe.
    pub(crate) fn slot(&self, chunk: &mut Chunk) -> u16 {
        chunk.alloc_scratch(1)
    }

    /// Push this function's own receiver. `local.get 0` under `Parameter`,
    /// the ambient global otherwise — the one place the difference is spelled.
    pub(crate) fn emit_receiver(&self, chunk: &mut Chunk, line: u32) {
        if self.declares_receiver {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        } else {
            crate::primitives::globals::emit_read(chunk, "__js_this", line);
        }
    }

    /// Emit `callee_slot(recv_slot)` — an invoke whose only argument is the
    /// receiver, which is how a zero-parameter method (`it.next()`,
    /// `v[Symbol.iterator]()`) is called.
    ///
    /// ⛔ The two ABIs need OPPOSITE emission orders, which is why this is a
    /// method and not a flag on the call site. Under `Ambient` the receiver is
    /// pushed FIRST because it feeds the global store, and the invoke takes no
    /// argument at all. Under `Parameter` the callee is pushed first and the
    /// receiver follows it as argument 0. Sharing a single `local.get` between
    /// them would strand a value on the stack in one of the two.
    ///
    /// ⛔ Unlike [`bind_ambient_receiver`], these sites pass argc 0 today — the
    /// receiver is NOT already on the argument list, so `Parameter` has to add
    /// it rather than just drop the ambient store.
    pub(crate) fn emit_receiver_invoke(
        &self,
        chunk: &mut Chunk,
        callee_slot: u16,
        recv_slot: u16,
        line: u32,
    ) {
        let argc = if self.abi == ReceiverAbi::Parameter {
            chunk.emit_op_u16(Op::LOCAL_GET, callee_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, recv_slot, line);
            1
        } else {
            chunk.emit_op_u16(Op::LOCAL_GET, recv_slot, line);
            crate::primitives::globals::emit_write(chunk, "__js_this", line);
            chunk.emit_op_u16(Op::LOCAL_GET, callee_slot, line);
            0
        };
        crate::primitives::callable::emit_direct_invoke_chunk(chunk, argc, line);
    }

    /// Save the CALLER's ambient receiver, so this helper's internal rebinds
    /// don't leak back out. Returns the slot holding it.
    ///
    /// ⛔ `None` under `Parameter`, and that is M5's whole point: with the
    /// receiver travelling as an argument there is no shared cell to clobber,
    /// so the save, the restore and the local they need all disappear. This is
    /// the hand-rolled shadow stack, and it is deleted rather than ported.
    ///
    /// Call it AFTER the other slots — the first `arity` slots are the
    /// parameters, and this one is never one of them.
    pub(crate) fn save_ambient_this(&self, chunk: &mut Chunk, line: u32) -> Option<u16> {
        if self.abi == ReceiverAbi::Parameter {
            return None;
        }
        let slot = self.slot(chunk);
        crate::primitives::globals::emit_read(chunk, "__js_this", line);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        Some(slot)
    }

    /// Undo [`Self::save_ambient_this`]. A `None` saved slot emits nothing.
    pub(crate) fn restore_ambient_this(&self, chunk: &mut Chunk, saved: Option<u16>, line: u32) {
        let Some(slot) = saved else { return };
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        crate::primitives::globals::emit_write(chunk, "__js_this", line);
    }

    /// Declare the finished frame. `arity`/`param_count` include the receiver
    /// where it is a REAL parameter — see the `param_count` note in
    /// `classes.rs`.
    ///
    /// ⛔ `local_count` is NOT assigned here. `alloc_scratch` has been
    /// maintaining it all along — for the frame's own slots and the emitter
    /// helpers' alike — so `finalize_local_count(0)` only re-asserts the
    /// `scratch_high_water` maximum. Assigning a count would clobber whatever
    /// those helpers allocated.
    ///
    /// ⛔ Under `Ambient` this writes `arity` and NOTHING else, reproducing
    /// exactly what the builders assigned before. `param_count` and
    /// `takes_receiver` stay at their defaults there, so the fourteen
    /// non-ambient languages see no change from this refactor.
    pub(crate) fn finish(&self, chunk: &mut Chunk, user_params: u8) {
        let receiver = u8::from(self.declares_receiver);
        chunk.arity = user_params + receiver;
        if receiver == 1 {
            chunk.param_count = chunk.arity;
            chunk.takes_receiver = true;
        }
        chunk.finalize_local_count(0);
    }
}

impl Compiler {
    pub(super) fn is_class_static_field_type_hint(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                // ⛔ ONE ANSWER: `Compiler::resolution_chain`. A hand-rolled
                // `parent` climb answers a DIAMOND by taking the first base's
                // chain and never seeing the others — the defect that made
                // `D().who()` return `'A'` where C3 says `'C'`.
                for cn in self.resolution_chain(&start) {
                    let Some(pc) = self.pending_classes.get(cn.as_str()) else {
                        continue;
                    };
                    let canon = self.canon(name);
                    if let Some(type_hint) = pc.static_field_types.get(&canon) {
                        return Some(type_hint.clone());
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }

    #[allow(dead_code)]
    pub(super) fn is_class_nested_type(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                for cn in self.resolution_chain(&start) {
                    let Some(pc) = self.pending_classes.get(cn.as_str()) else {
                        continue;
                    };
                    if pc.nested_types.iter().any(|n| {
                        if self.case_sensitive {
                            n == name
                        } else {
                            n.eq_ignore_ascii_case(name)
                        }
                    }) {
                        return Some(cn.to_string());
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }

    pub(super) fn generic_static_member_key(&self, type_expr: &str, field: &str) -> Option<String> {
        let expr = type_expr.trim();
        if !expr.contains('<') || !expr.contains('>') {
            return None;
        }

        let base = expr.split('<').next().map(str::trim).unwrap_or(expr);
        let base_canon = self.canon(base);
        if !self.defined_classes.contains(&base_canon) {
            return None;
        }

        let field_canon = self.canon(field);
        let has_static = self
            .pending_classes
            .get(base)
            .or_else(|| self.pending_classes.get(base_canon.as_str()))
            .map(|pc| pc.static_fields.iter().any(|f| f == &field_canon))
            .unwrap_or(false);
        if !has_static {
            return None;
        }

        let compact_type: String = expr.chars().filter(|c| !c.is_whitespace()).collect();
        let type_canon = self.canon(&compact_type);
        Some(format!("__gstatic_{}_{}", type_canon, field_canon))
    }

    /// Returns the owning class when `name` is a static method of the
    /// currently compiling class (or one of its ancestors).
    pub(super) fn is_class_static_method(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                for cn in self.resolution_chain(&start) {
                    let Some(pc) = self.pending_classes.get(cn.as_str()) else {
                        continue;
                    };
                    if pc.static_method_names.iter().any(|m| {
                        if self.case_sensitive {
                            m == name
                        } else {
                            m.eq_ignore_ascii_case(name)
                        }
                    }) {
                        return Some(cn.to_string());
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }

    pub(super) fn next_enclosing_class_name(&self, class_name: &str) -> Option<String> {
        self.pending_classes
            .get(class_name)
            .and_then(|pc| pc.enclosing_class.clone())
            .or_else(|| {
                class_name
                    .rsplit_once('.')
                    .map(|(outer, _)| outer.to_string())
            })
    }

    pub(super) fn class_extends_builtin(&self, class_name: &str, builtin: &str) -> bool {
        let mut current = Some(self.canon(class_name));
        let target = self.canon(builtin);
        while let Some(name) = current {
            let Some(pc) = self.pending_classes.get(name.as_str()) else {
                return false;
            };
            let Some(parent) = pc.parent.as_ref() else {
                return false;
            };
            let parent_canon = self.canon(parent);
            if parent_canon == target {
                return true;
            }
            current = Some(parent_canon);
        }
        false
    }

    /// Check if a name is a field of the current class (for implicit self resolution).
    pub(super) fn is_class_field(&self, name: &str) -> bool {
        if !self.current_class_implicit_self {
            return false;
        }
        self.current_class
            .as_deref()
            .and_then(|class_name| {
                self.visible_instance_field_storage_name_for_class(class_name, name)
            })
            .is_some()
    }

    /// Whether the current class (or an ancestor) declares `name` as an
    /// INSTANCE member of any kind — field, property, or method. This is the
    /// implicit-self question for members that are NOT plain fields: an
    /// accessor-backed property (`val y: Int` with a synthesized getter) has
    /// no entry in `fields`/`field_storage_names`, so `is_class_field` says
    /// no — yet a bare `y = …` in an init block is a write to `this.y` all
    /// the same, and must reach the property machinery, not a global.
    pub(super) fn is_class_instance_member(&self, name: &str) -> bool {
        if !self.current_class_implicit_self {
            return false;
        }
        let canon = self.canon(name);
        // ⛔ ONE ANSWER. This also carried its own `guard > 64` cycle cap —
        // `resolution_chain` is cycle-safe by construction, so the private
        // counter goes with the private walk.
        let Some(start) = self.current_class.as_deref().map(|c| self.canon(c)) else {
            return false;
        };
        for class_key in self.resolution_chain(&start) {
            let Some(pending) = self.pending_classes.get(class_key.as_str()) else {
                continue;
            };
            if pending.instance_member_names.iter().any(|m| m == &canon) {
                return true;
            }
        }
        false
    }

    pub(super) fn emit_self_ref(&mut self) -> bool {
        let self_kw = self.profile.self_keyword.clone();
        if let Some(self_slot) = self.scope().resolve(&self_kw) {
            self.emit_u16(Op::LOCAL_GET, self_slot);
            return true;
        }
        if self.scopes.len() > 1 {
            if let Some(_uv) = self.resolve_upvalue(self.scopes.len() - 1, &self_kw) {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index(&self_kw);
                let l = self.line;
                crate::primitives::closures::emit_env_get(self.chunk(), env, idx, l);
                return true;
            }
        }
        false
    }

    /// Profile-declared class dispatch model — `class_method_dispatch =
    /// "prototype"` in the language's `[compiler]` section. The shared
    /// class pipeline stays language-agnostic; languages opt in via the
    /// profile, never via name checks.
    pub(crate) fn class_prototype_dispatch(&self) -> bool {
        self.method_receiver_model() == Some(vybe_ast::MethodReceiver::Prototype)
    }

    /// Reading a method produces a fresh callable with the receiver already
    /// bound (Python). One of the three dispatch models.
    pub(crate) fn methods_bind_on_access(&self) -> bool {
        self.method_receiver_model() == Some(vybe_ast::MethodReceiver::BindOnAccess)
    }

    /// How a method call obtains its receiver, for this UNIT.
    ///
    /// The three models are mutually exclusive and were previously spread
    /// across three unrelated spellings — a profile string
    /// (`class_method_dispatch = "prototype"`), a profile bool
    /// (`methods_bind_on_access`), and a language NAME (`profile.name ==
    /// "php"`). One question, three answers, none of which could see the
    /// others; nothing prevented a profile from declaring two of them.
    ///
    /// Now one directive with three variants, stated by the walker on
    /// `Module.directives`, so it travels with the UNIT — a multi-language
    /// bundle answers per unit, which a profile installed once per compilation
    /// structurally cannot do.
    // `super_is_cooperative` is GONE. It read `profile.class_multiple_inheritance`
    // — a language-wide flag standing in for how a CALLEE is dispatched — and its
    // own docstring said so: "Still the WRONG CARRIER: by directives.md §3 this
    // describes how a CALLEE is dispatched, which is question 3 … which is the
    // most that can be fixed without the declaration-side field existing." The
    // field now exists: `NormalClass::cooperative_super`, read through
    // `classes_with_cooperative_super` where only the class NAME is in scope.

    /// A missing argument binds `undefined` (ECMA-262 §10.2.1.1).
    pub(crate) fn missing_arg_is_undefined(&self) -> bool {
        self.directives().missing_arg_is_undefined.unwrap_or(false)
    }

    /// Static fields are own properties of the class object.
    pub(crate) fn static_fields_are_own_properties(&self) -> bool {
        self.directives().static_fields_are_own_properties.unwrap_or(false)
    }

    /// In a SUBPROGRAM body, local declarations compile before nested
    /// procedures. See [`vybe_ast::Directives::body_declarations_first`].
    pub(crate) fn body_declarations_first(&self) -> bool {
        self.directives().body_declarations_first.unwrap_or(false)
    }

    /// Must a declared INSTANCE field be an own property of the instance —
    /// enumerable through the ordinary object surface? See
    /// [`vybe_ast::Directives::instance_fields_are_own_properties`].
    pub(crate) fn instance_fields_are_own_properties(&self) -> bool {
        self.directives()
            .instance_fields_are_own_properties
            .unwrap_or(false)
    }

    /// Is a declared function a first-class OBJECT carrying `name`, `length`,
    /// `prototype` and a `__nonenum` set — or just code? See
    /// [`vybe_ast::Directives::functions_are_objects`].
    ///
    /// Defaults to TRUE: the ECMA object model is what most of the wired
    /// walkers want, and a language whose functions are not objects says so.
    pub(crate) fn functions_are_objects(&self) -> bool {
        self.directives().functions_are_objects.unwrap_or(true)
    }

    /// Stamp `name` / `length` / `prototype` non-enumerably onto the function
    /// object on TOS, or DROP it when this unit's functions are not objects.
    ///
    /// A method rather than a bare call at each site so the directive read and
    /// the `self.chunk()` borrow do not collide, and so no site can forget it.
    pub(crate) fn stamp_fn_metadata_nonenum(&mut self, line: u32) {
        let objects = self.functions_are_objects();
        crate::primitives::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), objects, line);
    }

    /// Stamp the function object on TOS with its kind's intrinsic prototype,
    /// or DROP it when this unit's functions are not objects.
    pub(crate) fn stamp_function_kind_proto(&mut self, is_async: bool, is_generator: bool, line: u32) {
        let objects = self.functions_are_objects();
        crate::primitives::prototypes::emit_stamp_function_kind_proto(
            self.chunk(),
            objects,
            is_async,
            is_generator,
            line,
        );
    }

    /// Private members are internal slots, not properties (JS `#x`).
    pub(crate) fn supports_private_fields(&self) -> bool {
        self.profile.supports_private_fields
    }

    /// Properties and methods occupy separate namespaces.
    pub(crate) fn separate_property_method_namespace(&self) -> bool {
        self.profile.separate_property_method_namespace
    }

    // `class_introspection_metadata` is GONE — now
    // `NormalClass::introspection_metadata`, read directly off the class in
    // `compile_class` where both of its sites already had one in scope.

    /// Default argument expressions evaluate once at definition time.
    pub(crate) fn default_args_evaluated_once(&self) -> bool {
        self.profile.default_args_evaluated_once
    }

    /// ECMA `new` dispatch (§10.2.2): an explicitly returned object wins.
    pub(crate) fn ecma_new_dispatch(&self) -> bool {
        self.profile.ecma_new_dispatch
    }

    /// An `async` body is implicitly wrapped in try/catch.
    pub(crate) fn async_wraps_body_in_try(&self) -> bool {
        self.profile.async_wraps_body_in_try
    }

    /// Every function has an implicit `arguments` object.
    pub(crate) fn has_arguments_object(&self) -> bool {
        self.profile.has_arguments_object
    }

    /// Thrown errors carry the ECMA `Error` shape.
    pub(crate) fn ecma_error_object_shape(&self) -> bool {
        self.profile.ecma_error_object_shape
    }

    /// Methods are overridable without an explicit marker.
    pub(crate) fn methods_virtual_by_default(&self) -> bool {
        self.profile.methods_virtual_by_default
    }

    /// Wrong-arity calls are accepted rather than an error.
    pub(crate) fn relaxed_call_arity(&self) -> bool {
        self.profile.relaxed_call_arity
    }

    /// The language has `undefined` distinct from `null`.
    pub(crate) fn has_undefined_value(&self) -> bool {
        self.profile.has_undefined_value
    }

    // `class_member_metadata` is GONE. It asked whether a class's members carry
    // their declared metadata into the runtime — a property of the DECLARATION —
    // of a PROFILE, which is installed once per compilation and cannot answer
    // per unit. It is now `NormalClass::member_metadata`, set by the pascal
    // frontend the way php sets `late_static_binding`, and read directly where
    // the class is in scope.


    pub(crate) fn method_receiver_model(&self) -> Option<vybe_ast::MethodReceiver> {
        if let Some(model) = self.directives().method_receiver {
            return Some(model);
        }
        None
    }


    // `static_methods_take_receiver` is GONE. It answered "do STATIC methods
    // carry the called class as a receiver" with `profile.name == "php"` — a
    // language NAME standing in for a property of the DECLARATION. It is now
    // `NormalClass::late_static_binding`, set by the php frontend the way seven
    // frontends already set `explicit_self_param`, read directly where the
    // class is in scope and through `classes_with_late_static_binding` where
    // only its name is. One declaration, both ends.

    /// Must a method CALL pass the receiver as an explicit leading argument?
    ///
    /// Three models, and this predicate picks the third:
    /// - prototype dispatch (JS/Dart) rides `__js_this` and a bound-receiver
    ///   marker on the callable — see `class_prototype_dispatch`;
    /// - bind-on-access (Python) burns the receiver into a fresh bound method
    ///   when the method is READ — see `methods_bind_on_access`;
    /// - otherwise the callable is the raw function off the class struct and
    ///   carries no receiver, so the call site supplies one.
    ///
    /// NOT `explicit_method_receiver_argument`, which is Lua's and means the
    /// opposite — the walker ALREADY passed a receiver, so shared code must not
    /// add a second one.
    ///
    /// The declaration side already records this per callee as
    /// `chunk.is_method` (`classes.rs`, `has_receiver`); nothing reads it back,
    /// because `CallSignature` carries no receiver flag and a dynamic callee
    /// slot cannot reach its chunk. Thread it there and this predicate goes.
    pub(crate) fn call_supplies_receiver(&self) -> bool {
        // Stated by the walker on `Module.directives`, so it travels with the
        // UNIT and a multi-language bundle answers per unit. A profile is
        // installed once per compilation and cannot.
        if let Some(model) = self.directives().method_receiver {
            return model == vybe_ast::MethodReceiver::CallSite;
        }
        false
    }

    /// True for profiles whose comparison/equality operators dispatch to a
    /// user-defined dunder (`__eq__`/`__lt__`/… and their cross-language
    /// aliases) — i.e. the same profiles the `<`/`>` sites already route
    /// through `emit_rich_compare_locals` (Python, Ruby, Dart, C#, VB, …).
    /// Excludes JS (ECMA coercion), PHP (loose comparison) and Pascal.
    ///
    /// Dispatch goes through the `Eq`/`Lt`/`Compare` SLOTS: a language that
    /// binds them gets its own semantics, one that binds nothing falls back to
    /// primitive comparison. Declaring nothing IS the opt-out, so no language
    /// needs excluding here — pascal was, by name, and did not need to be.
    /// PHP stays out via `string_aware_relational`, which it declares anyway.
    pub(crate) fn uses_rich_comparison(&self) -> bool {
        !self.profile.ecma_operator_coercion && !self.profile.string_aware_relational
    }

    /// Operator overloading on the arithmetic/unary operators: a user
    /// `__add__`/`__neg__`/… on the operand wins over the primitive op.
    ///
    /// The same profiles that get rich comparison — a language either
    /// dispatches operators to methods or it coerces operands, and the
    /// two are the same question. Languages whose `+` is ECMA-coerced
    /// reach their operator methods through `ecma:value.add`'s
    /// ToPrimitive/`valueOf` chain instead.
    pub(crate) fn uses_rich_operators(&self) -> bool {
        self.uses_rich_comparison()
    }

    /// Box an i32 comparison result as a `Bool` — unless a CONDITION asked for
    /// the i32, in which case the boxing is skipped and that is reported.
    ///
    /// Skipping and reporting are deliberately the SAME statement. Two
    /// separate booleans would drift, and the two drift directions are not
    /// equally bad: skip-without-report only costs the ladder, but
    /// report-without-skip hands `BR_IF` a boxed `Bool` — which it accepts —
    /// and the loop branches on the wrong thing in total silence.
    pub(super) fn emit_i32_to_bool_or_report(&mut self) {
        if std::mem::take(&mut self.want_i32_condition) {
            self.gave_i32_condition = true;
            return;
        }
        let line = self.line;
        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
    }

    /// Compile `cond` and leave an **i32** 0/1 on the stack.
    ///
    /// The general path is `compile_expr` + `emit_condition_truthiness_from_stack`.
    /// But a relational operator has already produced an i32 — `emit_js_lt` and
    /// friends end in `f64.lt` — and the `emit_i32_to_bool` after them exists
    /// only for VALUE position. In condition position the truthiness ladder
    /// undid it immediately, via its own `js-boolean:test` + `js-boolean:cast`:
    /// three host calls and two branches to turn an i32 into an i32.
    ///
    /// Soundness rests on two things, neither of them a promise made here:
    /// `compile_expr` TAKES the request at entry, so `a < b && c < d` compiles
    /// its comparisons with the request clear and still boxes them; and only
    /// emitters whose result provably came from a WASM compare opcode honour
    /// it. `emit_rich_compare_locals` never does — its dunder arm returns the
    /// user's `__lt__` value, which can be any object — so Python and Pascal
    /// keep the full ladder.
    pub(super) fn compile_condition_to_i32(
        &mut self,
        cond: &vybe_ast::Expression,
    ) -> Result<(), String> {
        self.want_i32_condition = true;
        self.gave_i32_condition = false;
        let result = self.compile_expr(cond);
        self.want_i32_condition = false;
        let gave = std::mem::take(&mut self.gave_i32_condition);
        result?;
        if !gave {
            self.emit_condition_truthiness_from_stack();
        }
        Ok(())
    }

    /// Turn the value on the stack into an i32 truth, by the rule the
    /// `truthiness` DIRECTIVE states.
    ///
    /// This is the ONE place that answers "is this true". Every site that turns
    /// a value into a condition routes here — `if`, `while`, `and`/`or`,
    /// `bool()`, and `Unary{Not}`. They used to decide separately and drifted:
    /// `emit_dyn_not` never asked, so `assert []`, which desugars to a
    /// hand-built `Unary{Not}`, silently passed under a protocol language.
    ///
    /// Under [`Truthiness::Protocol`] the ladder is CPython §3.3.1 verbatim —
    /// [`ProtocolSlot::Bool`], then [`ProtocolSlot::Len`], then the value. A
    /// builtin `[]` is falsy through the SAME `Len` rung a user class with
    /// `__len__` uses, so there is no "empty collections" special case and a
    /// class in any language earns the behaviour by binding the slot.
    pub(super) fn emit_condition_truthiness_from_stack(&mut self) {
        if !self.protocol_truthiness() {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            return;
        }

        let line = self.line;
        let value_slot = self.define_local("__truth_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        // ── rung 0: absent is false ─────────────────────────────────────
        // Nothing below can run on a null: `ecma:value.typeof` traps on one,
        // so the very first probe took the whole ladder down and `if None:`
        // never reached the coercion that would have answered it.
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::REF_IS_NULL);
        self.chunk().emit_if_i32(line);
        inst!(self, core_wasm::i32_const, 0);
        self.chunk().emit_else(line);

        // STRUCT_GET traps on a primitive, so every slot probe sits behind the
        // object test — the same gate `emit_rich_unary` uses.
        let typeof_idx = self.import("ecma:value", "typeof");
        let is_object_slot = self.define_local("__truth_is_object");
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_host_call(typeof_idx, 1);
        self.emit_const(Value::String(Arc::from("object")));
        fn_call!(self, "wasm:js-string", "equals", 2);
        self.emit_u16(Op::LOCAL_SET, is_object_slot);

        // ── rung 1: ProtocolSlot::Bool ──────────────────────────────────
        let bool_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal(&vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Bool)));
        let bool_method = self.define_local("__truth_bool_method");
        self.emit_ref_null_local(bool_method);
        self.emit_u16(Op::LOCAL_GET, is_object_slot);
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &bool_key);
        self.emit_u16(Op::LOCAL_SET, bool_method);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, bool_method);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if_i32(line);
        self.emit_u16(Op::LOCAL_GET, bool_method);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::callable::emit_direct_invoke_chunk(self.chunk(), 1, line);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_else(line);

        // ── rung 2: ProtocolSlot::Len on a user class ───────────────────
        let len_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal(&vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Len)));
        let len_method = self.define_local("__truth_len_method");
        self.emit_ref_null_local(len_method);
        self.emit_u16(Op::LOCAL_GET, is_object_slot);
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.class_get_resolved(class_slots::ObjSource::Stack, &len_key);
        self.emit_u16(Op::LOCAL_SET, len_method);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, len_method);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if_i32(line);
        self.emit_u16(Op::LOCAL_GET, len_method);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::callable::emit_direct_invoke_chunk(self.chunk(), 1, line);
        fn_call!(self, "wasm:js-number", "toI32", 1);
        inst!(self, core_wasm::i32_const, 0);
        self.chunk().emit_op(Op::I32_NE, line);
        self.chunk().emit_else(line);

        // ── rung 3: the builtin length — a str/array/map answers `Len`
        //            intrinsically, so this is the same rung, not a case ──
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_host_call(typeof_idx, 1);
        self.emit_const(Value::String(Arc::from("string")));
        fn_call!(self, "wasm:js-string", "equals", 2);
        self.emit_u16(Op::LOCAL_GET, is_object_slot);
        self.chunk().emit_op(Op::I32_OR, line);
        self.chunk().emit_if_i32(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::collections::emit_len(&mut self.chunks, self.current, line);
        inst!(self, core_wasm::i32_const, 0);
        self.chunk().emit_op(Op::I32_NE, line);

        // ── rung 4: the value itself ────────────────────────────────────
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_end(line);

        self.chunk().emit_end(line); // rung 2 (Len slot)
        self.chunk().emit_end(line); // rung 1 (Bool slot)
        self.chunk().emit_end(line); // rung 0 (null)
    }

    /// `local = null` — a slot probe starts empty so the non-object path skips
    /// the `STRUCT_GET` without leaving the local undefined.
    fn emit_ref_null_local(&mut self, slot: u16) {
        let line = self.line;
        self.chunk()
            .emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        self.emit_u16(Op::LOCAL_SET, slot);
    }

}

/// How the receiver reaches the callee at ONE call site.
///
/// This used to be an `Option<u16>`, which could say *whether* a receiver
/// travels but not *how* — and M5 turns "how" into the whole question. The
/// three members of the protocol below read this one value, so a site
/// cannot save under one mechanism and restore under another.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ReceiverBind {
    /// No receiver protocol at a call in this region — the fourteen
    /// languages whose calls carry no `this` at all.
    None,
    /// The receiver is **argument 0** of the call — ECMA-262 §10.2.1
    /// `[[Call]](thisArgument, argumentsList)`. `slot` parks it between the
    /// point it is computed and the point the callee reference is on the
    /// stack, because `call_ref` wants `[callee, receiver, args…]` and the
    /// receiver is computed first.
    Argument { slot: u16 },
}

impl ReceiverBind {
    /// Does a receiver travel at all here? Replaces the old
    /// `Option::is_some`, which every call site used to ask.
    pub(super) fn is_active(self) -> bool {
        !matches!(self, ReceiverBind::None)
    }
}

impl Compiler {
    pub(super) fn begin_receiver_bind(&mut self, local_name: &str) -> ReceiverBind {
        if self.universal_receiver() {
            // Nothing to save: no global is being clobbered, so there is no
            // caller state to put back. The slot only parks the value.
            let slot = self
                .scope()
                .resolve(local_name)
                .unwrap_or_else(|| self.define_local(local_name));
            return ReceiverBind::Argument { slot };
        }
        // ⛔ NO AMBIENT ARM. It read `__js_this` into a save slot and returned
        // `Ambient { saved }`; nothing declares that binding any more, so the
        // whole save/bind/restore triple it anchored is gone with it.
        ReceiverBind::None
    }

    /// Clear the ambient receiver so a CONSTRUCTION allocates instead of
    /// adopting whatever `this` happens to be live.
    ///
    /// A constructor reads its receiver from `__js_this` and allocates only
    /// when that global is absent — `struct.new_default` sits in the `else` of
    /// a null test at the top of every `__<Class>_ctor_N`. Every other
    /// `save_js_this` site pairs the save with a `bind_js_this_from_local`; the
    /// `New` emit was the one that saved and restored without ever writing a
    /// value in between, so inside an instance method the constructor found the
    /// enclosing receiver, skipped the allocation, and wrote its fields into
    /// it. `One bump() => One(this.v + 1)` answered `identical(a, b) == true`;
    /// a constructor in a class-static field initializer shared one object for
    /// the same reason.
    ///
    /// `new` is unconditional: it always makes a fresh object, whatever the
    /// caller's context. Clearing states that, rather than relying on the
    /// caller happening to have no receiver.
    pub(super) fn clear_js_this(&mut self) {
        {
            return;
        }
        let line = self.line;
        self.chunk()
            .emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        self.emit_global_write("__js_this");
    }

    /// Write the ambient receiver from the value on the stack.
    ///
    /// **Emits unconditionally.** The `receiver_binding` directive is answered
    /// ONCE, by `save_js_this`, whose `Option` says whether this language has an
    /// ambient receiver at all; a site that got `None` must not compute a
    /// receiver, must not push one, and must not call this. There is no second
    /// decision here and no silent no-op to fall into.
    ///
    /// This used to decide for itself — `if !false { return; }` —
    /// while every caller pushed unconditionally:
    ///
    /// ```ignore
    /// let saved = self.save_js_this("__js_prev_this_member");
    /// self.emit_u16(Op::LOCAL_GET, obj_slot);   // ALWAYS pushed
    /// self.set_js_this_from_stack();            // popped only if Ambient
    /// ```
    ///
    /// `ReceiverBinding::Ambient` is declared by js and dart ONLY, so the other
    /// fourteen languages leaked one operand per member read, at ten sites. The
    /// stray is invisible while nothing live sits under it — a statement
    /// boundary truncates it — and fatal where something does: `W(self.v)`
    /// inside a method left the ctor ref buried and `CALL_REF` took the stray as
    /// the callee (`Not a function`), and `{'k': self.v}` compiled to
    /// `{<V object>: 5}` — the stray DISPLACED the key, silently.
    ///
    /// A guard that emits nothing is dead code for the language it fires on,
    /// and dead code that is *also* half a pair is how the two halves drifted.
    /// `restore_js_this(None)` reads the same `Option`, so all three members of
    /// the save/bind/restore triple now agree by construction.
    pub(super) fn bind_receiver_from_stack(&mut self, bind: ReceiverBind) {
        match bind {
            // The receiver is an ARGUMENT, so it must end up on the stack ABOVE
            // the callee reference — and it was computed BELOW it. Park it; the
            // call site pushes it back with `push_receiver_argument` once the
            // callee is in place.
            ReceiverBind::Argument { slot } => self.emit_u16(Op::LOCAL_SET, slot),
            // ⛔ UNREACHABLE, AND NOW IT EMITS NOTHING RATHER THAN A GLOBAL
            // WRITE. Every one of the sixteen callers gates this on
            // `bind.is_active()`, which is `!matches!(self, None)`, so a
            // `None` bind never arrives — verified by enumerating them, not
            // assumed. The write was kept only because it was what the
            // unconditional pre-M5 code did, and it is the last place
            // `__js_this` could be written by a language that never declared
            // an ambient receiver.
            //
            // Emitting nothing is safe for the same reason it is unreachable:
            // a caller that does not call this also did not push a receiver,
            // so there is no operand left stranded on the stack.
            ReceiverBind::None => {}
        }
    }

    /// Push the receiver as **argument 0** and report the extra argument count
    /// the call site must add to its `argc`.
    ///
    /// Call it after the callee reference is on the stack and before the
    /// declared arguments, which is the `[callee, thisArgument, args…]` order
    /// `call_ref` consumes and ECMA-262 §10.2.1 specifies. Returns `0` for
    /// every non-`Argument` binding, so a site can call it unconditionally and
    /// add the result — which is what keeps the argument count and the receiver
    /// push from ever being decided separately.
    /// **The** answer to "must this frame mark its receiver as captured so an
    /// inner closure can reach it, and through which channel".
    ///
    /// One function, one answer — the same rule as [`Self::emit_receiver_value`].
    /// An arrow binds `this` LEXICALLY (§10.2.11) and the capture works BY NAME
    /// through the shared env, so the ENCLOSING frame has to put its receiver
    /// in that env. Under the ambient receiver the name must first be created
    /// from the global; where the receiver is already a PARAMETER the local
    /// exists and only needs marking.
    ///
    /// ⛔ THIS DECISION EXISTED TWICE AND THE COPIES DIVERGED — which is the
    /// whole reason it is one function now. The class-method site grew both
    /// arms; the lambda site, which compiles OBJECT-LITERAL methods, kept only
    /// the ambient one. So under `ReceiverBinding::UniversalParameter` an
    /// object-literal method never marked its receiver captured, and the result
    /// was a LAYOUT DISAGREEMENT rather than a missing value: the frame sized
    /// its shared-env array without the receiver while the arrow inside still
    /// asked `closure_env_index(self_kw)`, which APPENDS when the name is
    /// absent — so the arrow read an index past the end and got null.
    ///
    /// It hid because with `this` as the ONLY capture the list came out empty,
    /// no env was built, and the receiver reached the arrow another way. ONE
    /// other captured name — any local, any global, sync or async — exposed it:
    ///
    /// ```text
    /// const o = { n: "N", m() {
    ///   const f = () => this.n;
    ///   const z = 1;
    ///   queueMicrotask(() => z);   // ← this line breaks the line above
    ///   return f();                // undefined; "N" without it
    /// }};
    /// ```
    ///
    /// `declares_own_receiver` is the caller's own fact about ITS frame — a
    /// static method and a plain closure have no receiver to mark — not a
    /// language trait; the channel question is answered here from the
    /// directives, never from a profile flag.
    pub(crate) fn capture_receiver_for_inner_closures(
        &mut self,
        closures_use_this: bool,
        declares_own_receiver: bool,
    ) {
        if !closures_use_this {
            return;
        }
        // ⛔ UNIVERSAL ARM ONLY — THE AMBIENT ARMS STAY AT THEIR SITES.
        //
        // Unifying the ambient arm as well cost **156 csharp regressions**
        // (1765 → 1898, measured by Rook): the three sites genuinely differed —
        // one created a SOURCE local under the self keyword, two created an
        // internal `__js_this`, and only one carried a parent-frame guard — and
        // collapsing them onto one spelling changed what `this` resolved to
        // inside accessors in every ambient language. An expression-bodied
        // property reading a field answered `0` instead of `12.57`.
        //
        // The genuinely MISSING piece was never the ambient half: it was that
        // the lambda site — which compiles OBJECT-LITERAL methods — had no
        // `UniversalParameter` arm at all. That is what this owns, and it is
        // js-only, so it is gated by a corpus I can actually run. Unifying the
        // ambient half needs csharp/vb/java/kotlin/php/python evidence that
        // nobody has yet; until someone does, three spellings that WORK beat
        // one spelling that is prettier and wrong.
        let self_kw = self.profile.self_keyword.clone();
        if declares_own_receiver && self.universal_receiver() {
            self.current_closure_captured_locals.insert(self_kw);
        }
    }

    pub(super) fn push_receiver_argument(&mut self, bind: ReceiverBind) -> u8 {
        match bind {
            ReceiverBind::Argument { slot } => {
                self.emit_u16(Op::LOCAL_GET, slot);
                1
            }
            _ => 0,
        }
    }

    /// **The** answer to "what is `this` here". One function, one answer.
    ///
    /// flexclassplan.md §"Two implementations of `what is this here`, and they
    /// disagreed" (item 2): `ExprKind::This` and `emit_js_current_this_value`
    /// were two independent resolvers, and they did not agree. Verified
    /// divergences, all three in the FALLBACK path — their ambient branches
    /// were condition-for-condition identical, which is why the bug presented
    /// as `super.m()` handing the callee an undefined receiver rather than as
    /// a general breakage:
    ///
    /// | | `ExprKind::This` | `emit_js_current_this_value` |
    /// |---|---|---|
    /// | receiver-is-a-parameter guard | yes | **none** |
    /// | local names tried | `self_kw`, `Self`, `self`, `this` | **`self_kw` only** |
    /// | derived-ctor TDZ (§9.1.1.3.4) | yes | **none** |
    ///
    /// The union is the first column throughout; the second contributed
    /// nothing the union needs.
    ///
    /// ⚠ This resolves the receiver as it is bound TODAY, which still includes
    /// the ambient `__js_this` global. That global is not a WASM concept —
    /// core wasm has no `this`, and a mutable module global standing in for a
    /// parameter needs the `save`/`restore` pair around every call, i.e. a
    /// hand-rolled shadow stack for something the substrate models natively.
    /// Collapsing the two resolvers is the PREREQUISITE for removing it: while
    /// two sites disagreed about what `this` is, nothing could safely change
    /// what it resolves TO.
    pub(super) fn emit_receiver_value(&mut self) {
        // ⛔ A CHUNK THAT TOOK THE RECEIVER AS A PARAMETER ALREADY HAS ONE —
        // the ambient global is not it.
        //
        // A property accessor is compiled as a chunk whose FIRST PARAMETER is
        // the receiver (`classes.rs`, `declare_receiver_first_accessor`). The
        // ambient branch fires first and reads `__js_this`, a global that
        // accessor chunk never set — so `this` inside an accessor answered null
        // while the same `this` in a method on the SAME object in the SAME run
        // was correct.
        //
        // Invisible from parsed source: every walker emits `Ident("this")` for
        // an accessor, which resolves the local and works. It bites only a
        // producer of SYNTHESIZED class AST — dart's `core_classes`, flutter's
        // adapter classes. There is no split to memorise; `This` means the
        // receiver.
        // ⛔ ASK THE CHUNK, NOT A CALL TAG. `declare_receiver_first_accessor`
        // sets `takes_receiver` and the tag together — they state the SAME
        // fact — and `takes_receiver` is the one that belongs on the function:
        // a funcref's type IS `[params] → [results]`, so "does argument 0 hold
        // a receiver" is a property of the SIGNATURE, not a note carried
        // beside it. A side tag is the same out-of-band channel as
        // `__bound_args` and the old per-call receiver arity: something a
        // stock engine cannot see, that the caller must consult to know what
        // the callee means.
        //
        // The tag mechanism is left in place but is no longer read here.
        let receiver_is_a_parameter = self.chunks[self.current].takes_receiver;

        // ▶▶ M5 STEP 1 — A BOUND RECEIVER OUTRANKS THE AMBIENT GLOBAL.
        //
        // This test used to sit BELOW the ambient branch, so `this` answered
        // from `__js_this` even in a chunk that had the receiver as a real
        // local. `receiver_is_a_parameter` patched the one case anyone had
        // caught (the accessor tag) by SUPPRESSING the ambient branch; every
        // other bound-receiver shape still lost to the global.
        //
        // Ordering it this way is the whole of "there must be a `this`
        // REFERENCE": where the language has bound one, that binding IS the
        // answer, and the module global is what remains only for the shapes
        // that have not been converted yet. M0 unified the two READERS; this is
        // the first step that changes what the reader RESOLVES TO.
        //
        // ⛔ Deliberately BEFORE the ambient branch and deliberately not
        // guarded by `receiver_is_a_parameter`: a local named by the self
        // keyword is a receiver whether or not a call tag advertised it. The
        // tag guard stays below only for chunks that have no such local.
        let self_kw = self.profile.self_keyword.clone();
        if let Some(slot) = self
            .scope()
            .resolve(&self_kw)
            .or_else(|| self.scope().resolve("Self"))
            .or_else(|| self.scope().resolve("self"))
            .or_else(|| self.scope().resolve("this"))
        {
            // §9.1.1.3.4: inside a derived constructor `this` is in TDZ until
            // super() runs — reading it while the slot is still null throws a
            // ReferenceError rather than handing back null.
            if self.js_derived_ctor_ctx == Some((self.current, slot)) {
                let l = self.line;
                crate::primitives::classes::emit_this_initialized_guard(self.chunk(), slot, l);
            }
            self.emit_u16(Op::LOCAL_GET, slot);
            return;
        }

        // ⛔ THE AMBIENT FALLBACK IS GONE. It read `__js_this` when no bound
        // receiver was in scope; no language declares
        // `ReceiverBinding::Ambient` any more, so it had no reachable caller.
        if self.scopes.len() > 1 {
            // Arrow function: capture `this` from the enclosing scope.
            if self.resolve_upvalue(self.scopes.len() - 1, &self_kw).is_some() {
                let env = self.closure_env_slot();
                let idx = self.closure_env_index(&self_kw);
                let l = self.line;
                crate::primitives::closures::emit_env_get(self.chunk(), env, idx, l);
                return;
            }
            // ⛔ The `__js_this` upvalue arm is gone with the ambient binding:
            // an arrow captures the receiver under the SELF KEYWORD above.
        }

        // ⛔ EVERY PATH OUT OF HERE MUST LEAVE EXACTLY ONE VALUE. This used to
        // be the `else` of an `ambient_this()` test whose other arm read the
        // global; the read is gone, the PUSH is not. Dropping it left the
        // receiver operand missing and the stack one short — measured as
        // `super.fetch()` inside an arrow returning the function object
        // instead of calling it, across 37 class/super/inheritance tests.
        self.emit_null();
    }

    /// End the binding.
    ///
    /// ⛔ THERE IS NOTHING LEFT TO UNDO, AND THAT IS THE POINT OF M5. Only the
    /// ambient protocol had caller state to restore: it clobbered a module
    /// global, so every bind needed a matching write-back — a hand-rolled
    /// shadow stack around every call, which wasm has no concept of. A receiver
    /// passed as ARGUMENT 0 was never global and clobbers nothing, so the
    /// restore half of the pair disappears with the global rather than being
    /// kept as a no-op for symmetry.
    ///
    /// Kept as a function because the call sites read as save/bind/restore and
    /// deleting one third of that shape at 20 sites is a separate change.
    pub(super) fn end_receiver_bind(&mut self, _bind: ReceiverBind) {}

    pub(super) fn save_js_new_target(&mut self, local_name: &str) -> Option<u16> {
        if !self.profile.ecma_new_dispatch {
            return None;
        }
        let slot = self
            .scope()
            .resolve(local_name)
            .unwrap_or_else(|| self.define_local(local_name));
        self.emit_global_read("__js_new_target");
        self.emit_u16(Op::LOCAL_SET, slot);
        Some(slot)
    }

    pub(super) fn set_js_new_target_from_stack(&mut self) {
        if !self.profile.ecma_new_dispatch {
            return;
        }
        self.emit_global_write("__js_new_target");
    }

    pub(super) fn restore_js_new_target(&mut self, slot: Option<u16>) {
        let Some(slot) = slot else {
            return;
        };
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_global_write("__js_new_target");
    }

    pub(super) fn set_js_new_target_undefined(&mut self) {
        if !self.profile.ecma_new_dispatch {
            return;
        }
        let line = self.line;
        common::expressions::emit_undefined(self.chunk(), line);
        self.emit_global_write("__js_new_target");
    }

    pub(super) fn flatten_member_chain(&self, expr: &Expression) -> Vec<String> {
        match &expr.kind {
            ExprKind::Ident(name) => Self::strip_global_namespace_prefix(name)
                .replace("::", ".")
                .split('.')
                .map(str::trim)
                .filter(|part| !part.is_empty())
                .map(ToString::to_string)
                .collect(),
            ExprKind::This => vec![self.profile.self_keyword.clone()],
            ExprKind::Super => vec![
                self.profile
                    .base_keyword
                    .clone()
                    .unwrap_or_else(|| "super".into()),
            ],
            ExprKind::Member { object, field, .. } => {
                let mut parts = self.flatten_member_chain(object);
                parts.push(field.clone());
                if parts
                    .first()
                    .is_some_and(|part| part.eq_ignore_ascii_case("global"))
                {
                    parts.remove(0);
                }
                parts
            }
            _ => Vec::new(),
        }
    }

    /// Extract plain expressions from Argument slice.
    #[allow(dead_code)]
    pub(super) fn arg_exprs(args: &[Argument]) -> Vec<&Expression> {
        args.iter().map(|a| &a.value).collect()
    }

    // ════════════════════════════════════════════════════════════════════════
    // Statement compilation
    // ════════════════════════════════════════════════════════════════════════
}
