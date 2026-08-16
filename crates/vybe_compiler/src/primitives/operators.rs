//! Operator, arithmetic-coercion, and compound-assignment compilation.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — same pattern as
//! `builtins.rs`/`calls.rs`. Methods are `pub(super)` so the core compile
//! paths in `mod.rs` and sibling files can reach them.

use super::*;

// Primitive fallbacks for `emit_rich_binop`, which takes the fallback as a
// `fn(&mut Chunk, u32)` so a slot and its primitive op stay one decision. The
// comparison and `Add` slots already had `ops::emit_dyn_*` to pass; the plain
// arithmetic ops have no dynamic form, so they are named here.
fn emit_f64_sub(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_SUB, line);
}
fn emit_f64_mul(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_MUL, line);
}
fn emit_f64_div(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_DIV, line);
}

impl Compiler {
    pub(super) fn emit_to_primitive(&mut self, hint: &str) {
        let value_slot = self.define_local("__to_primitive_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        inst!(self, recipes::is_object);
        let line = self.line;
        // `is_object` is `ref.test struct` — it already yields an i32, and
        // `Op::IF` takes an i32 (or a Bool) directly. Running
        // `emit_dyn_to_bool` on it re-ran the ~33-instruction ToBoolean ladder
        // over a value that was already 0/1, TWICE per relational operator
        // (`coerce_top_two_to_primitive`) and twice per `- * /`
        // (`coerce_top_two_to_number`).
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        let idx = self.import("ecma:value", "toPrimitive");
        self.emit_const(Value::String(Arc::from(hint)));
        self.emit_host_call(idx, 2);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.chunk().emit_end(line);
    }

    /// JS profile: coerce both top-of-stack operands via the
    /// ToPrimitive polyfill, then to_f64 via the VM's existing
    /// `Value::as_f64` once the operand is no longer an Object.
    /// Used for `-`, `*`, `/`. Passes hint="number" per ECMA §7.1.4
    /// step 1 (ToNumber unboxes Objects with hint=number first).
    pub(super) fn coerce_top_two_to_number(&mut self) {
        let t_b = self.define_local("__binop_b");
        self.emit_u16(Op::LOCAL_SET, t_b);
        // a on top → coerce
        self.emit_to_primitive("number");
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit_to_primitive("number");
    }

    /// Runtime-polymorphic numeric binary op for dynamically-typed
    /// languages. Stack `[a, b]` → `[result]`. At runtime, if BOTH operands
    /// are BigInt, calls `ecma:bigint.<bigint_fn>` (which returns a
    /// `Value::BigInt`, so the result stays BigInt-typed through a chain);
    /// otherwise runs the SAME number path the static route would
    /// (`coerce_top_two_to_number` — honouring `valueOf`/ToPrimitive — then
    /// the f64 op), so non-BigInt behaviour is byte-for-byte unchanged.
    pub(super) fn emit_js_dynamic_arith(&mut self, bigint_fn: &str, number_op: NumberArith) {
        let b_slot = self.define_local("__dynarith_b");
        self.emit_u16(Op::LOCAL_SET, b_slot);
        let a_slot = self.define_local("__dynarith_a");
        self.emit_u16(Op::LOCAL_SET, a_slot);

        let test_bi = self.import("wasm:js-bigint", "test");
        self.emit_u16(Op::LOCAL_GET, a_slot);
        self.emit_host_call(test_bi, 1);
        self.emit_u16(Op::LOCAL_GET, b_slot);
        self.emit_host_call(test_bi, 1);
        self.emit(Op::I32_AND);
        let line = self.line;
        self.chunk().emit_if_value(line);

        // both BigInt → ecma:bigint.<fn>(a, b)
        self.emit_u16(Op::LOCAL_GET, a_slot);
        self.emit_u16(Op::LOCAL_GET, b_slot);
        let bi = self.import("ecma:bigint", bigint_fn);
        self.emit_host_call(bi, 2);

        self.chunk().emit_else(line);
        // number path — identical to the static route.
        self.emit_u16(Op::LOCAL_GET, a_slot);
        self.emit_u16(Op::LOCAL_GET, b_slot);
        self.coerce_top_two_to_number();
        match number_op {
            NumberArith::Sub => self.emit(Op::F64_SUB),
            NumberArith::Mul => self.emit(Op::F64_MUL),
            NumberArith::Div => self.emit(Op::F64_DIV),
            NumberArith::Mod => {
                let l = self.line;
                common::math::emit_c_fmod(self.chunk(), l);
            }
        }
        let line = self.line;
        self.chunk().emit_end(line);
    }

    /// §13.15.2 strict-mode assignment: a false [[Set]] result throws
    /// TypeError. Consumes the Bool left by the strict proxy set dispatch.
    pub(super) fn emit_strict_set_failure_check(&mut self) -> Result<(), String> {
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);
        self.emit_const(Value::String(Arc::from(
            "'set' on proxy: trap returned falsish",
        )));
        self.emit_js_exception_ctor_from_message_value("TypeError")?;
        common::errors::emit_throw(self.chunk(), line);
        self.chunk().emit_end(line);
        Ok(())
    }

    /// `++`/`--` step. ECMA §13.4: ToNumeric keeps the operand's numeric
    /// type — a BigInt operand steps by 1n (result stays BigInt), anything
    /// else by Number 1. Profiles without BigInt keep the plain number
    /// path byte-for-byte.
    pub(crate) fn emit_step_by_one(&mut self, add: bool) {
        let line = self.line;
        // The language's OWN string step, if it declares one —
        // `[builtin_slots.string] inc/dec`, PHP's alphanumeric increment
        // ("a"++ is "b"). LANGUAGE table only: ECMA has no string step
        // (§13.4 is ToNumeric), so there is no platform default and the
        // runtime string test below only exists when a binding says so.
        let string_step = self
            .profile
            .builtin_slots
            .get(
                vybe_ast::builtin_slots::BuiltinType::String,
                if add {
                    vybe_ast::ProtocolSlot::Inc
                } else {
                    vybe_ast::ProtocolSlot::Dec
                },
            )
            .map(str::to_string);
        if let Some(target) = string_step {
            let slot = self.define_local("__step_v");
            self.emit_u16(Op::LOCAL_SET, slot);
            let test_str = self.import("wasm:js-string", "test");
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_host_call(test_str, 1);
            self.chunk().emit_if_value(line);
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_slot_target(&target, 1, line, "string step");
            self.chunk().emit_else(line);
            self.emit_u16(Op::LOCAL_GET, slot);
            common::bigint::emit_step(self.chunk(), add, line);
            self.chunk().emit_end(line);
            return;
        }
        if self.bigint_semantics() {
            let slot = self.define_local("__step_v");
            self.emit_u16(Op::LOCAL_SET, slot);
            let test_bi = self.import("wasm:js-bigint", "test");
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_host_call(test_bi, 1);
            let line = self.line;
            self.chunk().emit_if_value(line);
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_const(Value::F64(1.0));
            let bi = self.import("ecma:bigint", if add { "add" } else { "sub" });
            self.emit_host_call(bi, 2);
            self.chunk().emit_else(line);
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_number_step(add, line);
            self.chunk().emit_end(line);
        } else {
            let line = self.line;
            self.emit_number_step(add, line);
        }
    }

    /// The Number half of ECMA §13.4 — `Number::add(ToNumber(v), ±1𝔽)`.
    ///
    /// ⛔ `++` is NOT `v + 1`, and routing it through the `+` emitter was a
    /// SPEC-MAPPING error, not a coercion detail. §13.4.4.1 reads
    /// `oldValue = ToNumeric(? GetValue(lhs))`, then adds 1 **in the operand's
    /// own numeric type**. There is no String branch anywhere in it. The `+`
    /// operator is a different algorithm — §13.15.3
    /// ApplyStringOrNumericBinaryOperator — which calls ToPrimitive on both
    /// sides and **concatenates when either is a String**. Sending `++` there
    /// made `"5"++` yield `"51"` where the spec says `6`, and `"abc"++` yield
    /// `"abc1"` where the spec says NaN.
    ///
    /// △ `--` never had the bug: it was already a bare `F64_SUB`, which
    /// coerces numerically. It is routed here too so both directions read from
    /// the same clause rather than agreeing by accident.
    ///
    /// ✔ Correct is also cheaper: ToNumber then one `f64` op, with no
    /// ToPrimitive, no string test and no concat arm to emit.
    ///
    /// The BigInt half stays with the caller — whether a bigint arm exists at
    /// all is `bigint_semantics()`, and a profile without it must keep the
    /// plain number path byte-for-byte.
    fn emit_number_step(&mut self, add: bool, line: u32) {
        // ⛔ NOT `ecma:value.toNumber`, despite the name. That host fn ends in
        // `Value::F64(p.as_f64())` — the VM's own coercion, as its doc says —
        // which is not §7.1.4 StringToNumber: it answers NaN for `""` (spec:
        // +0), for `"0x10"` (spec: 16) and for `[]` (spec: +0). `ecma:number
        // Number` is the coercion `Number(v)` and unary `+` already use, and
        // it is the one that matches the spec on all three.
        let number = self.import("ecma:number", "Number");
        self.emit_host_call(number, 1);
        self.emit_const(Value::F64(1.0));
        self.emit(if add { Op::F64_ADD } else { Op::F64_SUB });
        let from_f64 = self.import("wasm:js-number", "fromF64");
        self.emit_host_call(from_f64, 1);
        let _ = line;
    }

    /// JS profile: ToPrimitive(hint=number) on both operands. Used
    /// before DYN_LT / DYN_GT / DYN_LE / DYN_GE so string-string lex
    /// compare and Date/valueOf-overriding instances both work.
    ///
    /// `left_known_primitive` / `right_known_primitive` say that operand is a
    /// numeric constant this emitter pushed itself, so `ref.test struct` — the
    /// whole of `emit_to_primitive` — is provably false for it. The
    /// `LOCAL_SET`/`LOCAL_GET` pair still runs either way, so the stack shape
    /// is identical with or without the skip.
    pub(super) fn coerce_top_two_to_primitive(
        &mut self,
        left_known_primitive: bool,
        right_known_primitive: bool,
    ) {
        let t_b = self.define_local("__cmpop_b");
        self.emit_u16(Op::LOCAL_SET, t_b);
        if !left_known_primitive {
            self.emit_to_primitive("number");
        }
        self.emit_u16(Op::LOCAL_GET, t_b);
        if !right_known_primitive {
            self.emit_to_primitive("number");
        }
    }

    /// The language's three-way `Compare` target for strings, if it declares
    /// one — `[builtin_slots.string] compare` (builtinslotplan.md §2a).
    ///
    /// Reads the LANGUAGE table only, deliberately not `get_or` with the
    /// platform default. The default is `common:str_compare`, a pure string
    /// compare that would mangle numeric operands; a language opts into
    /// string-aware relational behaviour by declaring it, and one that declares
    /// nothing keeps the plain dynamic comparison it has today.
    ///
    /// This one lookup replaces THREE mechanisms that all answered the same
    /// question depending on who was asking: the `string_aware_relational`
    /// profile bool, the name-keyed `LanguageHooks::relational_compare`
    fn string_compare3_target(&self) -> Option<String> {
        self.profile
            .builtin_slots
            .get(
                vybe_ast::builtin_slots::BuiltinType::String,
                vybe_ast::ProtocolSlot::Compare,
            )
            .map(str::to_string)
    }

    /// Emit the language's three-way compare: `[a, b]` → `i32` in `{-1, 0, 1}`.
    ///
    /// Panics on an emit-target shape it cannot emit. Emitting nothing would
    /// leave both operands stranded on the stack while the caller pushed a `0`
    /// to compare against — silent stack corruption traceable only to a typo in
    /// a profile. The `[builtin_slots.*]` parser is already fatal on a bad key;
    /// this is the same contract one layer down.
    fn emit_compare3(&mut self, target: &str) {
        let line = self.line;
        if let Some(name) = target.strip_prefix("common:") {
            self.emit_common(name, 2, line);
        } else if let Some(rest) = target.strip_prefix("host:") {
            // `host:<module>:<fn>` — the same shape the profile parser accepts.
            let (module, func) = rest.rsplit_once(':').unwrap_or_else(|| {
                panic!("[builtin_slots] compare target `{target}` is not `host:<module>:<fn>`")
            });
            let idx = self.import(module, func);
            self.emit_host_call(idx, 2);
        } else {
            panic!(
                "[builtin_slots] compare target `{target}` must be `common:…` or `host:…`; \
                 a three-way compare has no opcode or stdlib form"
            );
        }
    }

    /// The language's `Contains` target for sets, if it declares one —
    /// `[builtin_slots.set] contains` (builtinslotplan.md §2a).
    ///
    /// LANGUAGE table only, for the same reason as [`Self::string_compare3_target`]:
    /// the platform default for `Contains` is a collection membership test, while
    /// the `in` operator on most profiles here is a KEY check
    /// (`ecma:object.hasOwn`, ECMA-262 §13.10.1). Falling back to the default
    /// would silently convert every language's `in` from "is this a key" to "is
    /// this a value". A language that means membership declares it.
    fn set_contains_target(&self) -> Option<String> {
        self.profile
            .builtin_slots
            .get(
                vybe_ast::builtin_slots::BuiltinType::Set,
                vybe_ast::ProtocolSlot::Contains,
            )
            .map(str::to_string)
    }

    /// The language's `Contains` target for arrays — `[builtin_slots.array]
    /// contains`. Same LANGUAGE-only reasoning as [`Self::set_contains_target`]:
    /// the platform default is a value-membership test, while `in` on the ECMA
    /// profiles is a key test, so a language opts in by declaring.
    pub(super) fn array_contains_target(&self) -> Option<String> {
        self.profile
            .builtin_slots
            .get(
                vybe_ast::builtin_slots::BuiltinType::Array,
                vybe_ast::ProtocolSlot::Contains,
            )
            .map(str::to_string)
    }

    /// Emit a `[builtin_slots.*]` target against `argc` operands already on the
    /// stack. `what` names the binding for the panic message.
    ///
    /// Panics on a shape it cannot emit — emitting nothing would strand the
    /// operands on the stack, surfacing far from the profile typo that caused
    /// it. An unknown `common:` NAME is reported by `emit_common` itself.
    pub(super) fn emit_slot_target(&mut self, target: &str, argc: u8, line: u32, what: &str) {
        if let Some(name) = target.strip_prefix("common:") {
            // `Compiler::emit_common`, not the free `dispatch::emit_common`:
            // it also tries the import-needing flavour and, crucially, calls
            // `sync_scope_slots_with_chunk` afterwards, so scratch slots the
            // target allocated cannot collide with later locals.
            self.emit_common(name, argc, line);
        } else if let Some(rest) = target.strip_prefix("host:") {
            let (module, func) = rest.rsplit_once(':').unwrap_or_else(|| {
                panic!("[builtin_slots] {what} target `{target}` is not `host:<module>:<fn>`")
            });
            let idx = self.import(module, func);
            self.emit_host_call(idx, argc);
        } else {
            panic!("[builtin_slots] {what} target `{target}` must be `common:…` or `host:…`");
        }
    }

    /// Call a 2-argument `Contains` target on operands already on the stack in
    /// `(collection, value)` order.
    fn emit_contains_target_call(&mut self, target: &str) {
        let line = self.line;
        if let Some(name) = target.strip_prefix("common:") {
            self.emit_common(name, 2, line);
        } else if let Some(rest) = target.strip_prefix("host:") {
            let (module, func) = rest.rsplit_once(':').unwrap_or_else(|| {
                panic!("[builtin_slots] contains target `{target}` is not `host:<module>:<fn>`")
            });
            let idx = self.import(module, func);
            self.emit_host_call(idx, 2);
        } else {
            panic!("[builtin_slots] contains target `{target}` must be `common:…` or `host:…`");
        }
    }

    /// Membership test over two SLOTS — the shape the runtime probe needs,
    /// which holds its operands in locals so it can test the container and then
    /// reuse both in whichever leg wins.
    fn emit_contains_from_slots(&mut self, collection: u16, value: u16, target: &str) {
        self.emit_u16(Op::LOCAL_GET, collection);
        self.emit_u16(Op::LOCAL_GET, value);
        self.emit_contains_target_call(target);
    }

    /// The language's `==` target — `[builtin_slots.string] eq`.
    ///
    /// The OPERATOR's semantics, not a container rule: PHP's `==` coerces
    /// (`"1" == 1`), and that decision belongs to the language. Housed on the
    /// `string` row for the same reason `compare` is — that is where languages
    /// diverge, and the operand types are not statically known.
    ///
    /// LANGUAGE table only: the platform's `==` is `ops::emit_dyn_eq`, and a
    /// language that declares nothing keeps it.
    fn loose_eq_target(&self) -> Option<String> {
        self.loose_cmp_target(vybe_ast::ProtocolSlot::Eq)
    }

    /// `!=`'s own binding — `[builtin_slots.string] ne`.
    ///
    /// A SEPARATE target rather than `Eq` plus a logical not: the language's
    /// emitter materializes a real Bool, and negating its result with
    /// `emit_dyn_not` throws that away, so `$a != $b` printed a raw `0` instead
    /// of `bool(false)`. PHP's `emit_php_loose_eq` already takes a `negate`
    /// flag for exactly this.
    fn loose_ne_target(&self) -> Option<String> {
        self.loose_cmp_target(vybe_ast::ProtocolSlot::Ne)
    }

    fn loose_cmp_target(&self, slot: vybe_ast::ProtocolSlot) -> Option<String> {
        self.profile
            .builtin_slots
            .get(vybe_ast::builtin_slots::BuiltinType::String, slot)
            .map(str::to_string)
    }

    /// The language's STRUCTURAL equality target for composite built-ins —
    /// `[builtin_slots.array] eq`.
    ///
    /// Housed on the `array` row because that is the shape both implementations
    /// actually probe for: Python's is order-independent SET equality and
    /// Dart's is record/tuple equality, and a tuple IS an array here. §3g
    /// corrected this plan's earlier claim that `value_eq` was a STRING slot —
    /// neither implementation touches strings.
    ///
    /// LANGUAGE table only. The platform fallback for `==` is reference /
    /// primitive equality (`ops::emit_dyn_eq`), and a language that wants deep
    /// equality for its composites declares it.
    fn structural_eq_target(&self) -> Option<String> {
        self.profile
            .builtin_slots
            .get(
                vybe_ast::builtin_slots::BuiltinType::Array,
                vybe_ast::ProtocolSlot::Eq,
            )
            .map(str::to_string)
    }

    /// Emit a membership test for stack `[value, collection]` → bool.
    ///
    /// The target is called `(collection, value)` — receiver first, which is the
    /// order `ecma:object.hasOwn` already uses on the generic path below and the
    /// order `ecma:set.has` wants. The operand stack is the other way round, so
    /// this spills both and reloads them swapped.
    pub(super) fn emit_contains(&mut self, target: &str) {
        let t_collection = self.define_local("__contains_collection");
        let t_value = self.define_local("__contains_value");
        self.emit_u16(Op::LOCAL_SET, t_collection);
        self.emit_u16(Op::LOCAL_SET, t_value);
        self.emit_contains_from_slots(t_collection, t_value, target);
    }

    /// The language's `Contains` targets for the three container shapes a
    /// RUNTIME probe can tell apart. `Some` only when all three are declared.
    ///
    /// This is `builtinslotplan.md` §2c's runtime path, and Python is why it
    /// exists: §3b measured that idiomatic Python resolves NO receiver type
    /// statically, so `x in y` cannot pick a binding at compile time and has to
    /// test `y` at run time. The three answers are the language's own, and they
    /// genuinely differ from the platform defaults — Python's dict leg is
    /// `hasIn`, which walks the prototype chain, where the default for
    /// `Map`/`Contains` is own-only `dict.has`. Requiring all three to be
    /// declared rather than defaulting the gaps is what makes replacing the old
    /// `is_python_profile()` branch behaviour-preserving.
    fn contains_probe_targets(&self) -> Option<(String, String, String)> {
        use vybe_ast::ProtocolSlot::Contains;
        use vybe_ast::builtin_slots::BuiltinType as T;
        let slots = &self.profile.builtin_slots;
        Some((
            slots.get(T::String, Contains)?.to_string(),
            slots.get(T::Array, Contains)?.to_string(),
            slots.get(T::Map, Contains)?.to_string(),
        ))
    }

    /// `value in collection` where the collection's type is known only at run
    /// time: test the container, then dispatch to that type's `Contains`
    /// binding. String first, then array, then everything else.
    pub(super) fn emit_contains_probe_from_locals(
        &mut self,
        collection: u16,
        value: u16,
        targets: &(String, String, String),
    ) {
        let (string_target, array_target, other_target) = targets.clone();
        let line = self.line;

        self.emit_u16(Op::LOCAL_GET, collection);
        fn_call!(self, "wasm:js-string", "test", 1);
        self.chunk().emit_if_value(line);
        self.emit_contains_from_slots(collection, value, &string_target);
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, collection);
        let is_array = self.import("ecma:array", "isArray");
        self.chunk().emit_call(is_array, 1, line);
        inst!(self, core_wasm::i32_const, 0);
        crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
        let array_line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), array_line);
        self.chunk().emit_if_value(array_line);
        self.emit_contains_from_slots(collection, value, &array_target);
        self.chunk().emit_else(array_line);
        self.emit_contains_from_slots(collection, value, &other_target);
        self.chunk().emit_end(array_line);
        self.chunk().emit_end(line);
    }

    /// `a <op> b` derived from the sign of the three-way compare — §2f: bind
    /// `Compare`, and `Lt`/`Le`/`Gt`/`Ge` follow from it rather than each
    /// carrying its own copy of the language's comparison rules.
    fn emit_relational_from_compare3(&mut self, target: &str, cmp_fn: fn(&mut Chunk, u32)) {
        self.emit_compare3(target);
        self.emit_const(Value::I32(0));
        let line = self.line;
        cmp_fn(self.chunk(), line);
        if self.profile.materialize_bool_results {
            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
        }
    }

    /// JS profile: ToPrimitive(hint=default) on both operands. Used
    /// before DYN_ADD per ECMA §13.15.4 — the `+` operator picks the
    /// "default" hint, which gives valueOf the first shot and falls
    /// back to toString.
    pub(super) fn coerce_top_two_to_default_primitive(&mut self) {
        let t_b = self.define_local("__addop_b");
        self.emit_u16(Op::LOCAL_SET, t_b);
        self.emit_to_primitive("default");
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit_to_primitive("default");
    }

    pub(super) fn emit_js_add_string_concat_from_locals(&mut self, lhs_slot: u16, rhs_slot: u16) {
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_to_primitive("string");
        self.emit_const(Value::String(Arc::from("")));
        common::strings::emit_str_concat(self.chunk(), line);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit_to_primitive("string");
        self.emit_const(Value::String(Arc::from("")));
        common::strings::emit_str_concat(self.chunk(), line);
        fn_call!(self, "wasm:js-string", "concat", 2);
    }

    pub(super) fn emit_js_add_numeric_from_locals(&mut self, lhs_slot: u16, rhs_slot: u16) {
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_const(Value::F64(f64::NAN));
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let rhs_line = self.line;
        self.chunk().emit_if_value(rhs_line);
        self.emit_const(Value::F64(f64::NAN));
        self.chunk().emit_else(rhs_line);
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };
        self.chunk().emit_end(rhs_line);
        self.chunk().emit_end(line);
    }

    #[allow(dead_code)]
    pub(super) fn emit_js_add(&mut self) {
        let rhs_slot = self.define_local("__js_add_rhs");
        let lhs_slot = self.define_local("__js_add_lhs");
        self.emit_u16(Op::LOCAL_SET, rhs_slot);
        self.emit_u16(Op::LOCAL_SET, lhs_slot);

        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        self.emit_const(Value::String(Arc::from("string")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);
        self.emit_js_add_string_concat_from_locals(lhs_slot, rhs_slot);
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        self.emit_const(Value::String(Arc::from("string")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        };
        let rhs_line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), rhs_line);
        self.chunk().emit_if_value(rhs_line);
        self.emit_js_add_string_concat_from_locals(lhs_slot, rhs_slot);
        self.chunk().emit_else(rhs_line);
        self.emit_js_add_numeric_from_locals(lhs_slot, rhs_slot);
        self.chunk().emit_end(rhs_line);
        self.chunk().emit_end(line);
    }

    /// Try a user operator method on the left operand, else `fallback`.
    /// Stack: `[lhs, rhs]` → `[result]`.
    ///
    /// Operator overloading is a method call the operands opt into, so the
    /// dispatch lives here next to every other operator lowering rather
    /// than in any one language's adapter — a class that defines
    /// `operator +` (Dart), `__add__` (Python) or `op_Addition` (C#)
    /// normalises to the same bound method name, so this reaches all of
    /// them, including across languages.
    /// Try the user operator bound to `slot` on the operands, else `fallback`.
    ///
    /// Takes a SLOT, not a spelling, so an operator cannot be dispatched by
    /// name here even by accident — and the same code reaches Python's
    /// `__add__`, Dart's `operator +` and C#'s `operator+` because all three
    /// fill `ProtocolSlot::Add`.
    ///
    /// `Sub`/`Mul`/`Div` reach it too — this used to be wired for `Add` alone,
    /// so a C# `operator -` compiled to a bare `F64_SUB` that never asked the
    /// class, which is what `uses_rich_operators` promised in its own doc.
    fn emit_rich_binop(&mut self, slot: vybe_ast::ProtocolSlot, fallback: fn(&mut Chunk, u32)) {
        let line = self.line;
        let rhs_slot = self.define_local("__rich_op_rhs");
        let lhs_slot = self.define_local("__rich_op_lhs");
        self.emit_u16(Op::LOCAL_SET, rhs_slot);
        self.emit_u16(Op::LOCAL_SET, lhs_slot);
        common::expressions::emit_rich_arithmetic(
            self.chunk(),
            lhs_slot,
            rhs_slot,
            &vybe_ast::protocol_slot_key(slot),
            fallback,
            line,
        );
    }

    /// Try a user unary-operator method on the operand, else `fallback`.
    /// Stack: `[operand]` → `[result]`.
    pub(super) fn emit_rich_unary(
        &mut self,
        protocol_slot: vybe_ast::ProtocolSlot,
        fallback: fn(&mut Chunk, u32),
    ) {
        let line = self.line;
        let slot = self.define_local("__rich_op_operand");
        self.emit_u16(Op::LOCAL_SET, slot);
        common::expressions::emit_rich_unary(
            self.chunk(),
            slot,
            &vybe_ast::protocol_slot_key(protocol_slot),
            fallback,
            line,
        );
    }

    /// A numeric literal the COMPILER ITSELF materialised — `f64.const 13`.
    ///
    /// Not inference and not a type hint: the operand IS a literal in the AST,
    /// so the value on the stack is a constant this emitter just wrote. Every
    /// runtime question the ECMA coercion path asks about it therefore has a
    /// compile-time answer.
    ///
    /// `Literal::BigInt` is a SEPARATE variant and deliberately excluded — a
    /// bigint is not f64-representable and must keep the `wasm:js-bigint` arm.
    fn is_emitted_number_literal(expr: &Expression) -> bool {
        matches!(
            expr.kind,
            ExprKind::Lit(Literal::Int(_)) | ExprKind::Lit(Literal::Float(_))
        )
    }

    /// Whether this operand is provably a NUMBER at runtime — the fact
    /// `emit_rich_compare_locals`' `left_is_number` parameter needs.
    ///
    /// ⚠ **A WEAKER claim than [`Self::is_emitted_number_literal`], and the two
    /// must never be swapped.** That one means "this emitter just wrote an
    /// `f64.const`", which is what licenses skipping `toF64` and comparing with
    /// a bare `f64.lt`. This one says only that the VALUE is a number — it may
    /// well be boxed — so every unboxing step has to stay. Feeding this into
    /// `a_known_num`/`b_known_num` would drop the unbox and let `as_i32`/
    /// `as_f64` coerce a boxed operand SILENTLY.
    ///
    /// The whole safety condition is `TypeBinding::Converting`. A
    /// `Descriptive` hint guarantees nothing about the next assignment, and
    /// both of the cases that reach here are Descriptive by construction: a
    /// python annotation (`x: int = "s"` runs) and the spelling
    /// `compile_var_decl` infers from an untyped initializer, where
    /// `let i = 0; i = "x"` is legal. So this fires on genuinely DECLARED
    /// types — pascal `var i: Integer`, C# `int i` — and nowhere else.
    fn expr_is_provably_number(&self, expr: &Expression) -> bool {
        if Self::is_emitted_number_literal(expr) {
            return true;
        }
        let ExprKind::Ident(name) = &expr.kind else {
            return false;
        };
        let Some(declared) = self.lookup_var_declared(name) else {
            return false;
        };
        if declared.binding != vybe_ast::TypeBinding::Converting {
            return false;
        }
        // A DECLARED spelling can name a user type where a literal never
        // could, and the builtin table answers by spelling alone — a Delphi
        // `type Decimal = record … class operator LessThan` classifies as
        // Double and would have its overload skipped in silence. The program's
        // own declaration outranks the platform table, exactly as it does in
        // `compile_var_decl`'s `As Object` refinement.
        if self
            .resolve_pending_class_name_for_type_hint(declared.spelling())
            .is_some()
        {
            return false;
        }
        self.hint_is_builtin_number(declared.spelling())
    }

    pub(super) fn compile_binop(&mut self, op: &BinOp) {
        self.compile_binop_operands(op, None);
    }

    /// `compile_binop`, plus the operand expressions when the caller has them.
    ///
    /// Same shape as the `&&`/`||` arm in `expressions.rs`, which already reads
    /// per-operand static facts to pick an operator's lowering: the decision
    /// belongs to the EXPRESSION, and this is the only way the fact reaches the
    /// operator without a per-language flag. `None` — every synthetic
    /// `compile_binop` call — lowers exactly as before.
    pub(super) fn compile_binop_operands(
        &mut self,
        op: &BinOp,
        operands: Option<(&Expression, &Expression)>,
    ) {
        let (left_num_lit, right_num_lit) = match operands {
            Some((left, right)) => (
                Self::is_emitted_number_literal(left),
                Self::is_emitted_number_literal(right),
            ),
            None => (false, false),
        };
        // The WEAKER of the two facts — "is a number", not "is a constant this
        // emitter just wrote". Only the rich-compare probe skip may read it;
        // see `expr_is_provably_number`.
        let left_is_number = match operands {
            Some((left, _)) => self.expr_is_provably_number(left),
            None => false,
        };
        match op {
            BinOp::Add => {
                // `dynamic_add`: JS-style `+` — concatenates when either
                // operand is a string, otherwise adds numerically. PHP,
                // Python, Lua, etc. use `.` / `..` / other operators for
                // string concat, so `+` is purely numeric and coerces
                // string operands (`"2026" + 4 == 2030`). `F64_ADD`
                // coerces both sides via `Value::as_f64()`; `DYN_ADD`
                // has the JS-style string-concat special case.
                if self.profile.dynamic_add {
                    // JS profile: ECMA §13.15.4 — call ToPrimitive on
                    // both operands with hint "default" before adding.
                    // The polyfill returns the operand unchanged for
                    // primitives (fast path) and unboxes Objects via
                    // their valueOf/toString chain (Date, custom
                    // valueOf, class instances).
                    if self.profile.ecma_operator_coercion {
                        let idx = self.import("ecma:value", "add");
                        self.emit_host_call(idx, 2);
                        return;
                    }
                    // A user `operator +` / `__add__` defines `+` for its own
                    // type; `emit_dyn_add` would coerce the operand instead
                    // and never consult it. Falls through to the same
                    // dynamic add for every non-object operand.
                    if self.uses_rich_operators() {
                        self.emit_rich_binop(
                            vybe_ast::ProtocolSlot::Add,
                            crate::primitives::ops::emit_dyn_add,
                        );
                        return;
                    }
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                } else if let Some(target) = self
                    .profile
                    .builtin_slots
                    .get(
                        vybe_ast::builtin_slots::BuiltinType::Array,
                        vybe_ast::ProtocolSlot::Add,
                    )
                    .map(str::to_string)
                {
                    // A language whose `+` is overloaded on collections as well
                    // as numbers (PHP's array UNION) declares
                    // `[builtin_slots.array] add`. Was
                    // `LanguageHooks::arith_add`, looked up by language NAME
                    // (builtinslotplan.md §3c).
                    let line = self.line;
                    self.emit_slot_target(&target, 2, line, "array add");
                } else {
                    self.emit(Op::F64_ADD);
                }
            }
            BinOp::Sub => {
                if self.profile.dynamic_numeric_dispatch {
                    self.emit_js_dynamic_arith("sub", NumberArith::Sub);
                } else if self.uses_rich_operators() {
                    self.emit_rich_binop(vybe_ast::ProtocolSlot::Sub, emit_f64_sub);
                } else {
                    if self.profile.ecma_operator_coercion {
                        self.coerce_top_two_to_number();
                    }
                    self.emit(Op::F64_SUB);
                }
            }
            BinOp::Mul => {
                if self.profile.dynamic_numeric_dispatch {
                    self.emit_js_dynamic_arith("mul", NumberArith::Mul);
                } else if self.uses_rich_operators() {
                    self.emit_rich_binop(vybe_ast::ProtocolSlot::Mul, emit_f64_mul);
                } else {
                    if self.profile.ecma_operator_coercion {
                        self.coerce_top_two_to_number();
                    }
                    self.emit(Op::F64_MUL);
                }
            }
            BinOp::Div => {
                if self.profile.dynamic_numeric_dispatch {
                    self.emit_js_dynamic_arith("div", NumberArith::Div);
                } else if self.uses_rich_operators() {
                    self.emit_rich_binop(vybe_ast::ProtocolSlot::Div, emit_f64_div);
                } else {
                    if self.profile.ecma_operator_coercion {
                        self.coerce_top_two_to_number();
                    }
                    self.emit(Op::F64_DIV);
                }
            }
            BinOp::IDiv => {
                self.emit(Op::F64_DIV);
                let l = self.line;
                common::math::emit_trunc(self.chunk(), l);
            }
            BinOp::FloorDiv => {
                self.emit(Op::F64_DIV);
                let l = self.line;
                common::math::emit_floor(self.chunk(), l);
            }
            BinOp::Mod => {
                let l = self.line;
                // `[builtin_slots.int] mod` — was `is_python_profile()`,
                // LANGUAGE table only: the platform's `%` truncates, and a
                // language whose `%` floors says so.
                if let Some(target) = self
                    .profile
                    .builtin_slots
                    .get(
                        vybe_ast::builtin_slots::BuiltinType::Int,
                        vybe_ast::ProtocolSlot::Mod,
                    )
                    .map(str::to_string)
                {
                    if let Some(name) = target.strip_prefix("common:") {
                        self.emit_common(name, 2, l);
                    } else if let Some(rest) = target.strip_prefix("host:") {
                        let (module, func) = rest.rsplit_once(':').unwrap_or_else(|| {
                            panic!("[builtin_slots.int] mod `{target}` is not `host:<m>:<fn>`")
                        });
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, 2);
                    } else {
                        panic!("[builtin_slots.int] mod `{target}` must be `common:…` or `host:…`");
                    }
                } else if self.profile.dynamic_numeric_dispatch {
                    self.emit_js_dynamic_arith("rem", NumberArith::Mod);
                } else {
                    common::math::emit_c_fmod(self.chunk(), l);
                }
            }
            BinOp::Pow => {
                let l = self.line;
                common::math::emit_pow(self.chunk(), l);
            }
            BinOp::Eq => {
                // `[builtin_slots.string] eq` — the language's own `==`.
                if let Some(target) = self.loose_eq_target() {
                    let line = self.line;
                    self.emit_slot_target(&target, 2, line, "string eq");
                    return;
                }
                if self.profile.abstract_equality {
                    let idx = self.import("ecma:value", "abstractEq");
                    self.emit_host_call(idx, 2);
                } else if self.uses_rich_comparison() {
                    // Dispatch to a user `__eq__` (or cross-language alias) with
                    // the receiver, falling back to structural equality.
                    let right_slot = self.define_local("__rich_eq_rhs");
                    let left_slot = self.define_local("__rich_eq_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    // Structural-equality fallback is the language's registered
                    // `value_eq` hook (present iff the language wants deep value
                    // equality — Python tuples/dicts, Dart records), else plain
                    // reference/primitive equality. Hook presence is the property;
                    // never a profile-name check.
                    // The language's STRUCTURAL equality for composite
                    // built-ins, from `[builtin_slots.array] eq`. Was
                    // `registry::hooks(&self.profile.name).value_eq` — a
                    // callback looked up by language NAME (§3c).
                    let eq_target = self.structural_eq_target();
                    common::expressions::emit_rich_compare_locals(
                        &mut self.chunks,
                        self.current,
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq),
                        match &eq_target {
                            Some(t) => common::expressions::RichFallback::Target(t),
                            None => common::expressions::RichFallback::Op(
                                crate::primitives::ops::emit_dyn_eq,
                            ),
                        },
                        line,
                        // Equality probes a different name set — type 0 DOES
                        // register `equals` — so that is its own decision.
                        false,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                } else {
                    {
                        let line = self.line;
                        // Value equality, as DECLARED. Both sides carrying the
                        // `__value_eq` stamp means two languages independently
                        // said their `==` compares fields — so compare fields,
                        // whoever allocated the objects. This is the read side
                        // of the policy; Dart and dotnet each had a private
                        // reader of the same stamp, which is what made a record
                        // lose its equality the moment it crossed a boundary.
                        let right_slot = self.define_local("__veq_rhs");
                        let left_slot = self.define_local("__veq_lhs");
                        self.emit_u16(Op::LOCAL_SET, right_slot);
                        self.emit_u16(Op::LOCAL_SET, left_slot);
                        crate::primitives::records::emit_is_value_eq(self.chunk(), left_slot, line);
                        crate::primitives::records::emit_is_value_eq(
                            self.chunk(),
                            right_slot,
                            line,
                        );
                        self.chunk().emit_op(Op::I32_AND, line);
                        self.chunk().emit_if(line);
                        crate::primitives::records::emit_value_fields_equal(
                            &mut self.chunks,
                            self.current,
                            left_slot,
                            right_slot,
                            line,
                        );
                        if !self.profile.materialize_bool_results {
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        }
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, left_slot);
                        self.emit_u16(Op::LOCAL_GET, right_slot);
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                        if self.profile.materialize_bool_results {
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                        }
                        self.chunk().emit_end(line);
                    };
                }
            }
            BinOp::NotEq => {
                // `[builtin_slots.string] ne` — its OWN target, see there.
                if let Some(target) = self.loose_ne_target() {
                    let line = self.line;
                    self.emit_slot_target(&target, 2, line, "string ne");
                    return;
                }
                if self.profile.abstract_equality {
                    let idx = self.import("ecma:value", "abstractNe");
                    self.emit_host_call(idx, 2);
                } else if self.uses_rich_comparison() {
                    // `a != b` == not (a `__eq__` b) — dispatch `__eq__`, negate.
                    let right_slot = self.define_local("__rich_ne_rhs");
                    let left_slot = self.define_local("__rich_ne_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    // Structural-equality fallback is the language's registered
                    // `value_eq` hook (present iff the language wants deep value
                    // equality — Python tuples/dicts, Dart records), else plain
                    // reference/primitive equality. Hook presence is the property;
                    // never a profile-name check.
                    // The language's STRUCTURAL equality for composite
                    // built-ins, from `[builtin_slots.array] eq`. Was
                    // `registry::hooks(&self.profile.name).value_eq` — a
                    // callback looked up by language NAME (§3c).
                    let eq_target = self.structural_eq_target();
                    common::expressions::emit_rich_compare_locals(
                        &mut self.chunks,
                        self.current,
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq),
                        match &eq_target {
                            Some(t) => common::expressions::RichFallback::Target(t),
                            None => common::expressions::RichFallback::Op(
                                crate::primitives::ops::emit_dyn_eq,
                            ),
                        },
                        line,
                        // Equality probes a different name set — type 0 DOES
                        // register `equals` — so that is its own decision.
                        false,
                    );
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                } else {
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                        if self.profile.materialize_bool_results {
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                        }
                    };
                }
            }
            BinOp::StrictEq => {
                let line = self.line;
                if self.profile.ecma_operator_coercion {
                    crate::primitives::ops::emit_js_strict_eq(self.chunk(), line);
                } else {
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                }
                // A comparison RESULT is a boolean, normalized here rather than
                // left to a per-language flag. The VM has no true/false, so the
                // raw comparison is an i32 and this is what makes it a value.
                // php set neither `ecma_boolean_operators` nor
                // `materialize_bool_results`, so `var_dump(1 === 2)` printed `0`
                // where real php prints `bool(false)` — while `==` was already
                // correct, because php binds it through
                // `[builtin_slots.string] eq` and its own emitter returns a
                // bool. Languages keep their quirks in the GRAMMAR (php spells
                // two equalities); the result TYPE is normalized.
                crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
            }
            BinOp::StrictNotEq => {
                // JS !==: negate of ===.
                {
                    let line = self.line;
                    if self.profile.ecma_operator_coercion {
                        crate::primitives::ops::emit_js_strict_eq(self.chunk(), line);
                        self.emit(Op::I32_EQZ);
                    } else {
                        crate::primitives::ops::emit_dyn_ne(self.chunk(), line);
                    }
                    // Same normalization as `StrictEq` above.
                    crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                }
            }
            BinOp::Lt => {
                // builtinslotplan.md §2f — the language's three-way `Compare`
                // decides, and this operator is its sign. Checked BEFORE
                // `ecma_operator_coercion` per §2d: a language's own binding
                // outranks a platform-wide coercion default.
                if let Some(target) = self.string_compare3_target() {
                    self.emit_relational_from_compare3(
                        &target,
                        crate::primitives::ops::emit_dyn_lt,
                    );
                    return;
                }
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive(left_num_lit, right_num_lit);
                }
                if self.profile.ecma_operator_coercion {
                    // ECMA-262 §7.2.13 — NaN-safe ToNumber on mixed operands.
                    let line = self.line;
                    crate::primitives::ops::emit_js_lt(
                        self.chunk(),
                        line,
                        left_num_lit,
                        right_num_lit,
                    );
                    // Ends in `f64.lt`, so the result IS an i32 — a condition
                    // can take it as-is instead of having it boxed and then
                    // immediately unboxed by the ToBoolean ladder.
                    self.emit_i32_to_bool_or_report();
                } else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(
                        &mut self.chunks,
                        self.current,
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Lt),
                        common::expressions::RichFallback::Op(crate::primitives::ops::emit_dyn_lt),
                        line,
                        left_is_number,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::Gt => {
                // builtinslotplan.md §2f — the language's three-way `Compare`
                // decides, and this operator is its sign. Checked BEFORE
                // `ecma_operator_coercion` per §2d: a language's own binding
                // outranks a platform-wide coercion default.
                if let Some(target) = self.string_compare3_target() {
                    self.emit_relational_from_compare3(
                        &target,
                        crate::primitives::ops::emit_dyn_gt,
                    );
                    return;
                }
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive(left_num_lit, right_num_lit);
                }
                if self.profile.ecma_operator_coercion {
                    // ECMA-262 §7.2.13 — NaN-safe ToNumber on mixed operands.
                    let line = self.line;
                    crate::primitives::ops::emit_js_gt(
                        self.chunk(),
                        line,
                        left_num_lit,
                        right_num_lit,
                    );
                    self.emit_i32_to_bool_or_report();
                } else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(
                        &mut self.chunks,
                        self.current,
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Gt),
                        common::expressions::RichFallback::Op(crate::primitives::ops::emit_dyn_gt),
                        line,
                        left_is_number,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::LtEq => {
                // builtinslotplan.md §2f — the language's three-way `Compare`
                // decides, and this operator is its sign. Checked BEFORE
                // `ecma_operator_coercion` per §2d: a language's own binding
                // outranks a platform-wide coercion default.
                if let Some(target) = self.string_compare3_target() {
                    self.emit_relational_from_compare3(
                        &target,
                        crate::primitives::ops::emit_dyn_le,
                    );
                    return;
                }
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive(left_num_lit, right_num_lit);
                }
                if self.profile.ecma_operator_coercion {
                    // ECMA-262 §7.2.13 — NaN-safe ToNumber on mixed operands.
                    let line = self.line;
                    crate::primitives::ops::emit_js_le(
                        self.chunk(),
                        line,
                        left_num_lit,
                        right_num_lit,
                    );
                    self.emit_i32_to_bool_or_report();
                } else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(
                        &mut self.chunks,
                        self.current,
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Le),
                        common::expressions::RichFallback::Op(crate::primitives::ops::emit_dyn_le),
                        line,
                        left_is_number,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::GtEq => {
                // builtinslotplan.md §2f — the language's three-way `Compare`
                // decides, and this operator is its sign. Checked BEFORE
                // `ecma_operator_coercion` per §2d: a language's own binding
                // outranks a platform-wide coercion default.
                if let Some(target) = self.string_compare3_target() {
                    self.emit_relational_from_compare3(
                        &target,
                        crate::primitives::ops::emit_dyn_ge,
                    );
                    return;
                }
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_primitive(left_num_lit, right_num_lit);
                }
                if self.profile.ecma_operator_coercion {
                    // ECMA-262 §7.2.13 — NaN-safe ToNumber on mixed operands.
                    let line = self.line;
                    crate::primitives::ops::emit_js_ge(
                        self.chunk(),
                        line,
                        left_num_lit,
                        right_num_lit,
                    );
                    self.emit_i32_to_bool_or_report();
                } else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(
                        &mut self.chunks,
                        self.current,
                        left_slot,
                        right_slot,
                        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Ge),
                        common::expressions::RichFallback::Op(crate::primitives::ops::emit_dyn_ge),
                        line,
                        left_is_number,
                    );
                    if self.profile.materialize_bool_results {
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    }
                }
            }
            BinOp::Spaceship => {
                // `<=>` is ITS OWN operator, not three others stacked up.
                // When the language declares a three-way `Compare`, emit it
                // directly — one call, and the operands are evaluated once.
                if let Some(target) = self.string_compare3_target() {
                    self.emit_compare3(&target);
                    return;
                }
                // No declared `Compare`: derive the sign from `<` then `>`.
                // Two evaluations of each operand, which is why a language that
                // cares declares the primitive instead.
                let right_slot = self.define_local("__spaceship_rhs");
                let left_slot = self.define_local("__spaceship_lhs");
                self.emit_u16(Op::LOCAL_SET, right_slot);
                self.emit_u16(Op::LOCAL_SET, left_slot);

                self.emit_u16(Op::LOCAL_GET, left_slot);
                self.emit_u16(Op::LOCAL_GET, right_slot);
                if self.profile.ecma_operator_coercion {
                    // `<=>` reloads the operands from slots, so the top two are
                    // still (left, right) and the literal facts still hold.
                    self.coerce_top_two_to_primitive(left_num_lit, right_num_lit);
                } else {
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                }
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                self.emit_const(Value::I32(-1));
                self.chunk().emit_else(line);

                self.emit_u16(Op::LOCAL_GET, left_slot);
                self.emit_u16(Op::LOCAL_GET, right_slot);
                if self.profile.ecma_operator_coercion {
                    // `<=>` reloads the operands from slots, so the top two are
                    // still (left, right) and the literal facts still hold.
                    self.coerce_top_two_to_primitive(left_num_lit, right_num_lit);
                } else {
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
                    };
                }
                let gt_line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), gt_line);
                self.chunk().emit_if_value(gt_line);
                self.emit_const(Value::I32(1));
                self.chunk().emit_else(gt_line);
                self.emit_const(Value::I32(0));
                self.chunk().emit_end(gt_line);
                self.chunk().emit_end(line);
            }
            BinOp::And | BinOp::Or => unreachable!(), // handled with short-circuit
            BinOp::Xor => {
                self.emit(Op::I32_XOR);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
            }
            BinOp::Eqv => {
                self.emit(Op::I32_XOR);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            }
            BinOp::Imp => {
                let rhs_slot = self.define_local("__imp_rhs");
                let lhs_slot = self.define_local("__imp_lhs");
                self.emit_u16(Op::LOCAL_SET, rhs_slot);
                self.emit_u16(Op::LOCAL_SET, lhs_slot);
                self.emit_u16(Op::LOCAL_GET, lhs_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
                self.emit_u16(Op::LOCAL_GET, rhs_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.emit(Op::I32_OR);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
            }
            BinOp::BitAnd => self.emit(Op::I32_AND),
            BinOp::BitOr => self.emit(Op::I32_OR),
            BinOp::BitXor => self.emit(Op::I32_XOR),
            BinOp::Shl => self.emit(Op::I32_SHL),
            BinOp::Shr => self.emit(Op::I32_SHR_S),
            BinOp::UShr => {
                self.emit(Op::I32_SHR_U);
                if self.profile.ecma_operator_coercion {
                    // ECMA-262 §13.10.2: `>>>` produces an unsigned 32-bit
                    // integer (Number). I32_SHR_U leaves the bit pattern
                    // in i32, but if the high bit is set (e.g. `-1 >>> 0`)
                    // the i32 → Number coercion would render as negative.
                    // Reinterpret as u32 by adding 2^32 when the i32 is
                    // negative — keeps within the f64 53-bit mantissa.
                    self.emit(Op::F64_FROM_I32);
                    let unsigned_slot = self.define_local("__js_ushr_result");
                    self.emit_u16(Op::LOCAL_SET, unsigned_slot);
                    self.emit_u16(Op::LOCAL_GET, unsigned_slot);
                    self.emit_const(Value::F64(0.0));
                    self.emit(Op::F64_LT);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, unsigned_slot);
                    self.emit_const(Value::F64(4_294_967_296.0));
                    self.emit(Op::F64_ADD);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, unsigned_slot);
                    self.chunk().emit_end(line);
                }
            }
            BinOp::Concat => {
                let l = self.line;
                common::strings::emit_str_concat_coercing(self.chunk(), l);
            }
            BinOp::In => {
                // §2c RUNTIME path: the language declares a `Contains` for
                // each container shape and the probe picks one at run time.
                if let Some(targets) = self.contains_probe_targets() {
                    let t_y = self.define_local("__in_probe_y");
                    let t_x = self.define_local("__in_probe_x");
                    self.emit_u16(Op::LOCAL_SET, t_y);
                    self.emit_u16(Op::LOCAL_SET, t_x);
                    self.emit_contains_probe_from_locals(t_y, t_x, &targets);
                    return;
                }

                // A language declares `[builtin_slots.set] contains`; the
                // shared path emits it without knowing the language.
                if let Some(target) = self.set_contains_target() {
                    self.emit_contains(&target);
                    return;
                }

                // `x in y` — JS: is `x` a property KEY of `y` (not a value).
                // ECMA-262 §13.10.1 walks the prototype chain. PHP
                // `in_array` / Python `key in dict` are own-only and
                // route through their language profiles separately.
                //
                // Walker stack: `[x, y]`. hasIn expects `[y, x]`.
                let l = self.line;
                let t_y = self.define_local("__in_y");
                let t_x = self.define_local("__in_x");
                self.emit_u16(Op::LOCAL_SET, t_y);
                self.emit_u16(Op::LOCAL_SET, t_x);

                // Proxy has-trap dispatch on the JS profile when the
                // module references `Proxy`. Stack: [obj, key].
                if self.uses_proxy {
                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    self.emit_proxy_has()
                        .expect("proxy has hook must exist when proxy lowering is enabled");
                    return;
                }

                self.emit_u16(Op::LOCAL_GET, t_y);
                self.emit_u16(Op::LOCAL_GET, t_x);
                // JS uses prototype-walking `hasIn`; other languages
                // (case-insensitive profiles or non-JS) keep own-only
                // `hasOwn` semantics for their `in`-shaped operators.
                let import = if self.profile.ecma_in_operator {
                    "hasIn"
                } else {
                    "hasOwn"
                };
                let idx = self.import("ecma:object", import);
                self.chunk().emit_call(idx, 2, l);
                // hasIn/hasOwn return Value::Bool — already correct for ECMA display.
            }
            BinOp::NotIn => {
                // The negation of `In`'s runtime probe — see there.
                if let Some(targets) = self.contains_probe_targets() {
                    let t_y = self.define_local("__nin_probe_y");
                    let t_x = self.define_local("__nin_probe_x");
                    self.emit_u16(Op::LOCAL_SET, t_y);
                    self.emit_u16(Op::LOCAL_SET, t_x);
                    self.emit_contains_probe_from_locals(t_y, t_x, &targets);
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    return;
                }

                // The negation of `In`'s slot binding — see there.
                if let Some(target) = self.set_contains_target() {
                    self.emit_contains(&target);
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                    return;
                }

                let l = self.line;
                let t_y = self.define_local("__nin_y");
                let t_x = self.define_local("__nin_x");
                self.emit_u16(Op::LOCAL_SET, t_y);
                self.emit_u16(Op::LOCAL_SET, t_x);
                self.emit_u16(Op::LOCAL_GET, t_y);
                self.emit_u16(Op::LOCAL_GET, t_x);
                // Same key-check as `in` above — route through hasOwn.
                let idx = self.import("ecma:object", "hasOwn");
                self.chunk().emit_call(idx, 2, l);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            }
            BinOp::InstanceOf => {
                if self.class_prototype_dispatch() {
                    let rhs_slot = self.define_local("__js_instanceof_rhs");
                    let lhs_slot = self.define_local("__js_instanceof_lhs");
                    self.emit_u16(Op::LOCAL_SET, rhs_slot);
                    self.emit_u16(Op::LOCAL_SET, lhs_slot);
                    // ECMA-262 §13.10.2: `a instanceof B` first checks for
                    // `B[Symbol.hasInstance]` (canonical name `hasinstance`)
                    // and calls it as `B[hasinstance](a)` if present.
                    // Compiler-side dispatch keeps the JS method-call
                    // protocol intact (`__js_this` bound to B) — host
                    // `ctx.invoke` can't do that, so we emit the
                    // method-call inline instead of going through the
                    // host fn for this case.
                    let has_inst_key = self.str_const("hasinstance");
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, has_inst_key);
                    let method_slot = self.define_local("__has_inst_method");
                    self.emit_u16(Op::LOCAL_SET, method_slot);
                    self.emit_u16(Op::LOCAL_GET, method_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    let helper = self.import("ecma:value", "instanceOf");
                    self.emit_u16(Op::LOCAL_GET, lhs_slot);
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.emit_host_call(helper, 2);
                    self.chunk().emit_else(line);
                    let saved_this = self.save_js_this("__js_prev_this_hasinst");
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, method_slot);
                    self.emit_u16(Op::LOCAL_GET, lhs_slot);
                    self.emit_direct_callable_invoke(1);
                    {
                        let line = self.line;
                        // Convert dynamic result to Bool (consistent with
                        // instanceOf host fn which also returns Bool).
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                    };
                    let result_slot = self.define_local("__has_inst_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.restore_js_this(saved_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.chunk().emit_end(line);
                } else {
                    // Dynamic-RHS fallback: the static `a instanceof TypeName`
                    // form is intercepted upstream in `expressions.rs` and
                    // emitted as `Op::REF_TEST` directly. This branch only
                    // fires for the rare `a instanceof <expression>` shape.
                    //
                    // Stack on entry: [val, ctor]. We string-compare
                    // `val.__type` against `ctor.name` — the same compile-time
                    // type-stamp the constructors install via `set_type_id`.
                    let l = self.line;
                    // Stack [val, ctor]. Extract the constructor's NAME string,
                    // then membership-check it against val's `__types` ancestry
                    // (falling back to the single `__type`). Inheritance-aware
                    // and shared: `is_a?`/`isinstance` normalize to this node.
                    let t_ctor = self.define_local("__io_ctor");
                    self.emit_u16(Op::LOCAL_SET, t_ctor); // [val]
                    self.emit_u16(Op::LOCAL_GET, t_ctor);
                    let name_key = self.str_const("name");
                    self.chunk()
                        .emit_struct_field_op(Op::STRUCT_GET, 0, name_key, l); // [val, ctor_name]
                    common::reflection::emit_instanceof(&mut self.chunks, self.current, l);
                }
            }
            BinOp::NullCoalesce => unreachable!(), // handled in compile_expr
            BinOp::MatMul => {
                let i = self.import("ecma:math", "matmul");
                self.emit_host_call(i, 2);
            }
            BinOp::Like => {
                // The VB walker always rewrites `a Like b` to `Regex.IsMatch(pattern, a)`
                // before reaching this point, so this arm is dead for VB.
                // Stack at call site: [string, pattern]; ecma:regexp.test expects
                // (pattern, string) — callers via the walker path never reach here.
                let idx = self.import("ecma:regexp", "test");
                self.emit_host_call(idx, 2);
            }
            BinOp::Is => {
                // Reference equality
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
            }
            BinOp::IsNot => {
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Compound assignment operator emission
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_compound_op(&mut self, op: &CompoundOp) {
        match op {
            CompoundOp::Add => {
                if self.profile.dynamic_add {
                    if self.profile.ecma_operator_coercion {
                        self.coerce_top_two_to_default_primitive();
                    }
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
                    };
                } else {
                    self.emit(Op::F64_ADD);
                }
            }
            CompoundOp::Sub => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_SUB);
            }
            CompoundOp::Mul => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_MUL);
            }
            CompoundOp::Div => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_DIV);
            }
            CompoundOp::IDiv => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_DIV);
                let l = self.line;
                common::math::emit_trunc(self.chunk(), l);
            }
            CompoundOp::Mod => {
                if self.profile.ecma_operator_coercion {
                    self.coerce_top_two_to_number();
                }
                let l = self.line;
                common::math::emit_c_fmod(self.chunk(), l);
            }
            CompoundOp::Pow => {
                let l = self.line;
                common::math::emit_pow(self.chunk(), l);
            }
            CompoundOp::Concat => {
                let l = self.line;
                common::strings::emit_str_concat(self.chunk(), l);
            }
            CompoundOp::BitAnd => self.emit(Op::I32_AND),
            CompoundOp::BitOr => self.emit(Op::I32_OR),
            CompoundOp::BitXor => self.emit(Op::I32_XOR),
            CompoundOp::Shl => self.emit(Op::I32_SHL),
            CompoundOp::Shr => self.emit(Op::I32_SHR_S),
            CompoundOp::UShr => self.emit(Op::I32_SHR_U),
            CompoundOp::And => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            } // simplified
            CompoundOp::Or => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            } // simplified
            CompoundOp::NullCoalesce => {
                // a ??= b → if a is null, a = b
                // At this point both are on stack already — no-op, the whole compound assign handles it
            }
        }
    }

    pub(super) fn is_csharp_delegate_handler_expr(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lambda { .. } | ExprKind::FuncRef(_) | ExprKind::CallableRef { .. } => true,
            ExprKind::Ident(name) => {
                if self
                    .lookup_var_type_hint(name)
                    .is_some_and(Self::is_callable_type_hint)
                {
                    return true;
                }
                if self.scope().resolve(name).is_some() {
                    return false;
                }
                let cname = self.canon(name);
                self.defined_functions.contains(&cname)
                    || self.defined_class_methods.contains(&cname)
            }
            ExprKind::Member { field, .. } => {
                let cname = self.canon(field);
                self.defined_functions.contains(&cname)
                    || self.defined_class_methods.contains(&cname)
            }
            ExprKind::New { args, .. } if args.len() == 1 => {
                self.is_csharp_delegate_handler_expr(&args[0].value)
            }
            _ => false,
        }
    }

    pub(super) fn assign_target_matches_expr(
        &self,
        target: &Expression,
        expr: &Expression,
    ) -> bool {
        match (&target.kind, &expr.kind) {
            (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
            (
                ExprKind::Member {
                    object: to,
                    field: tf,
                    ..
                },
                ExprKind::Member {
                    object: eo,
                    field: ef,
                    ..
                },
            ) => {
                if !tf.eq_ignore_ascii_case(ef) {
                    return false;
                }
                self.assign_target_matches_expr(to, eo)
            }
            (ExprKind::This, ExprKind::This) => true,
            _ => false,
        }
    }

    // ════════════════════════════════════════════════════════════════════════
}
