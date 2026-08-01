//! The PLATFORM DEFAULT table — `builtinslotplan.md` step 2.
//!
//! What a built-in type does for a protocol slot when the language declares
//! nothing. §2e: the default is a decision **on record**, not an emergent
//! property of which shared helper a profile happened to point at.
//!
//! Nothing reads this yet, by design. Steps 1–4 establish that the table
//! faithfully describes current behaviour BEFORE any of it changes, because
//! the alternative — flip and see what breaks across twelve languages — is not
//! measurable (§6).
//!
//! # Why ECMA is the default
//!
//! Measured 2026-07-30 (`builtinslotplan.md` §3a): across 9 string slots in
//! php/python/js/dart, **32 of 36 cells already agree with the real runtime**,
//! and every divergence is PHP. So ECMA behaviour as the default means eleven
//! of twelve languages declare nothing and stay correct, and PHP declares the
//! handful of places it genuinely differs:
//!
//! | slot | php | python / js / dart |
//! |---|---|---|
//! | `Len` `"café"` | 5 (bytes) | 4 (UTF-16 units) |
//! | `Eq` `"1e2" == "100"` | true | false |
//! | `Bool` `"0"` | falsy | truthy |
//! | `Add` | `.` concatenates, `+` is arithmetic | `+` concatenates |
//!
//! # A caution about the existing `str_*` surface
//!
//! Each entry below names the target whose measured behaviour matches, not the
//! one whose name matches.
//!
//! **Corrected 2026-08-01.** This note previously claimed `common:str_length`
//! (`strings::emit_length` → `wasm:js-string.length`) counts **UTF-8 bytes**,
//! and that the `str_*` surface was therefore split between byte and UTF-16
//! semantics. That is FALSE, read from the registration sites:
//!
//! - `wasm:js-string.length` — `s.encode_utf16().count()` as `I32`
//!   (`crates/vybe_runtime/src/js_string_builtins.rs`)
//! - `ecma:string.length` — `s.encode_utf16().count()` as `F64`
//!   (`platforms/ecma/src/string.rs`)
//!
//! Both are **UTF-16 code units**; they differ only in return type. The `str_*`
//! surface is unit-CONSISTENT, and the default below is uniform rather than
//! carefully picked around a split that does not exist.
//!
//! The real consequence is the opposite of what the old note implied: the
//! platform had **no shared byte-length primitive at all**. PHP built its own
//! (`php::string_adapter::emit_strlen`, a UTF-16 walk summing UTF-8 widths),
//! and it was the only one — so a language that wanted bytes had nothing to
//! bind to. See `unifiedstringplan.md` §1a.
//!
//! **Resolved 2026-08-01.** That counter moved into
//! `strings::emit_byte_length` and is now reachable as `common:str_byte_length`,
//! so all three units are bindable targets:
//!
//! | unit | target | who counts this way |
//! |---|---|---|
//! | UTF-16 code units | `host:ecma:string:length` (the default below) | js, java, dart, c#, vb |
//! | code points | `common:str_scalar_length` | python `len`, php `mb_strlen` |
//! | UTF-8 bytes | `common:str_byte_length` | php `strlen`, lua `#`, go `len(s)` |
//!
//! Nothing binds the byte target yet — php's `strlen` calls the emitter
//! directly because it also applies php's own coercion first. The point is that
//! the unit is now a DECLARABLE value, which is what §3a asked for.

use crate::primitives::Compiler;
use std::sync::OnceLock;
use vybe_ast::builtin_slots::{BuiltinSlotBindings, BuiltinType};
use vybe_ast::{Expression, ProtocolSlot};

/// The platform default for every `(BuiltinType, ProtocolSlot)` the platform
/// can emit today.
///
/// A pair absent from this table is **unbound**, which is a real state, not an
/// oversight — see [`unbound_reason`].
pub fn platform_defaults() -> BuiltinSlotBindings {
    let mut b = BuiltinSlotBindings::new();

    // ── string ──────────────────────────────────────────────────────────
    // UTF-16 code units: ECMA-262 `String.prototype.length`. NOT
    // `common:str_length`, which is the byte count.
    b.insert(
        BuiltinType::String,
        ProtocolSlot::Len,
        "host:ecma:string:length",
    );
    // `s[i]` yields a one-character string. Verified against Dart 2026-07-30:
    // `"café"[3]` is `é`; routing this through `ecma:array.get` returned null
    // for every index.
    //
    // CAVEAT, same shape as `Map`/`GetItem`: languages disagree OUT OF RANGE
    // (JS `undefined`, Dart `RangeError`, Python `IndexError`). Bound because
    // in-range indexing was verified to agree; a language whose out-of-range
    // answer differs declares `[builtin_slots.string] get_item`. Read an
    // out-of-range failure as this caveat, not as a new finding.
    b.insert(
        BuiltinType::String,
        ProtocolSlot::GetItem,
        "common:str_char_at",
    );
    b.insert(
        BuiltinType::String,
        ProtocolSlot::Contains,
        "common:str_contains",
    );
    b.insert(BuiltinType::String, ProtocolSlot::Add, "common:str_concat");
    b.insert(
        BuiltinType::String,
        ProtocolSlot::Compare,
        "common:str_compare",
    );
    // The platform answer for `==` / `!=` IS `ops::emit_dyn_eq` — that is what
    // `BinOp::Eq` falls back to when no language declares otherwise. Recording
    // it makes the default a decision on record (§2e) instead of an emergent
    // property of one `else` branch, and removes the old claim that "no emit
    // target exists" — `common:dyn_eq` always had one.
    //
    // PHP overrides both: its `==` coerces (`"1" == 1`).
    b.insert(BuiltinType::String, ProtocolSlot::Eq, "common:dyn_eq");
    b.insert(BuiltinType::String, ProtocolSlot::Ne, "common:dyn_ne");

    // ── array ───────────────────────────────────────────────────────────
    b.insert(
        BuiltinType::Array,
        ProtocolSlot::Len,
        "common:collections.length",
    );
    // Same out-of-range caveat as `String`/`GetItem` above.
    b.insert(
        BuiltinType::Array,
        ProtocolSlot::GetItem,
        "common:collections.get",
    );
    b.insert(
        BuiltinType::Array,
        ProtocolSlot::SetItem,
        "common:collections.set",
    );
    b.insert(
        BuiltinType::Array,
        ProtocolSlot::Contains,
        "common:collections.contains",
    );
    b.insert(
        BuiltinType::Array,
        ProtocolSlot::Add,
        "common:collections.concat",
    );

    // ── map ─────────────────────────────────────────────────────────────
    b.insert(BuiltinType::Map, ProtocolSlot::Len, "common:dict.size");
    // The ECMA answer: a missing key is `undefined`. Languages whose miss
    // differs — Dart's `null`, Python's `KeyError`, PHP's `null`+warning —
    // declare `[builtin_slots.map] get_item` and win by §2d precedence.
    //
    // This pair was UNBOUND from 2026-07-31 until the per-language override
    // table existed, because binding it centrally made Dart's `json['body']`
    // on an empty map return `undefined` where it must be `null`. That is the
    // whole reason overrides exist, so the two landed together.
    b.insert(
        BuiltinType::Map,
        ProtocolSlot::GetItem,
        "common:dict.get_dynamic",
    );

    // ── bytes ───────────────────────────────────────────────────────────
    // Indexing a byte string yields the INTEGER byte in Python, Go and PHP 8
    // alike, which is what `at` returns for a `Uint8Array`. The out-of-range
    // caveat above applies here too, and a language that differs declares
    // `[builtin_slots.bytes] get_item`.
    b.insert(
        BuiltinType::Bytes,
        ProtocolSlot::GetItem,
        "host:ecma:uint8array:at",
    );
    b.insert(
        BuiltinType::Map,
        ProtocolSlot::SetItem,
        "common:dict.set_dynamic",
    );
    b.insert(BuiltinType::Map, ProtocolSlot::Contains, "common:dict.has");

    b
}

/// Why a `(type, slot)` pair has no default.
///
/// §5's definition of done requires every unread slot to have "a reader or a
/// written reason for having none". This is where the reasons live, so an
/// unbound pair can be told apart from a forgotten one.
pub fn unbound_reason(ty: BuiltinType, slot: ProtocolSlot) -> Option<&'static str> {
    use BuiltinType as T;
    use ProtocolSlot as S;
    Some(match (ty, slot) {
        // `(String, Eq)` and `(String, Ne)` are NOT here any more: PHP binds
        // both (`common:php.loose_eq` / `.loose_ne`) and the platform default
        // is `ops::emit_dyn_eq`. The old reason claimed no emit target existed;
        // measurement showed `common:dyn_eq` always had.
        (T::String, S::Bool) | (T::Array, S::Bool) | (T::Map, S::Bool) => {
            "Truthiness is emitted inline by `ops::emit_dyn_to_bool` rather than \
             through a dispatchable target. PHP is the only language that \
             differs (`\"0\"` is falsy), so this stays unbound until step 6 \
             gives it a target."
        }
        (T::String, S::Hash) => {
            "Must land WITH `Eq` (§2g): binding one without the other yields \
             values that compare equal and miss in maps."
        }
        (T::String, S::Iterator) => {
            "Languages disagree on what iterating a string yields — code points \
             (JS, Dart), characters (Python), runes with byte offsets (Go), \
             nothing at all (PHP). Needs measuring before a default is chosen; \
             guessing here would encode one language's answer as everyone's."
        }
        (T::String, S::Lt) | (T::String, S::Le) | (T::String, S::Gt) | (T::String, S::Ge) => {
            "Derived from `Compare`'s sign per flexclassplan §2c-bis — not \
             bound separately."
        }
        // `GetItem` IS bound above, so the blanket arm must not claim it —
        // the same contradiction `(Int, Mod)` had once Python bound it.
        (T::Bytes, S::GetItem) => return None,
        (T::Bytes, _) => {
            "`Literal::Bytes` now exists (unifiedstringplan.md §3c) and \
             `GetItem` is bound above; the remaining bytes slots are unmeasured, \
             not blocked."
        }
        // `(Int, Mod)` is bound by Python (floored `%`), so `Int` can no longer
        // claim a blanket reason — only the slots nobody has measured.
        (T::Int, S::Mod) => return None,
        (T::Int, _) | (T::Double, _) | (T::BigInt, _) => {
            "No target chosen yet. The numeric slots were left out of the \
             default table deliberately: §3a measured STRING behaviour across \
             four languages, and the numeric rows are still listed under \
             'Still to measure'. Binding them from the same reasoning would be \
             the guess §2e exists to prevent. Separately, they are also \
             unresolvable — see `unresolvable_reason`."
        }
        (T::Object, _) => {
            "A plain object's slots come from its class, and a user class always \
             wins over a built-in (§2d). Binding `object` would invert that."
        }
        _ => return None,
    })
}

/// Why [`Compiler::builtin_type_of`] can never NAME this type, independent of
/// whether its slots are bound.
///
/// A DIFFERENT axis from [`unbound_reason`], and conflating the two would
/// mislead exactly the audit §5 asks for. "No target exists" needs someone to
/// write an emitter; "the resolver cannot name this type" needs someone to fix
/// the resolver. An auditor reading a single `Some(reason)` for `array`/`Len`
/// would conclude it is handled, when what it lacks is resolution, not a
/// binding.
///
/// The two axes are ORTHOGONAL, and the numeric types show why: `int`, `double`
/// and `bigint` are unbound *and* unresolvable — nobody has chosen a target,
/// and nothing could reach it if they had. Requiring a binding here would force
/// those two independent facts into one answer.
pub fn unresolvable_reason(ty: BuiltinType) -> Option<&'static str> {
    use BuiltinType as T;
    Some(match ty {
        T::Int | T::Double | T::BigInt => {
            "The platform's only numeric classifier, \
             `Compiler::is_numeric_type_hint`, returns a bool over ~20 spellings \
             (`int32`, `longint`, `real`, `single`, `sbyte`, …) and cannot say \
             which of int/double/bigint a hint names. Adding that discriminator \
             here would be a new per-language spelling table in shared code — \
             the anti-pattern this plan retires. It arrives with step 4, when \
             each language declares its own numeric spellings."
        }
        T::Array => {
            "Not resolvable from a static hint: `expr_is_array_like` is a \
             VALUE-SHAPE heuristic (array literal, `lookup_array_binding`, \
             `array()`/`str_split()` calls, arithmetic on arrays), not a \
             static-type read. Resolving `Array` off it would capture receivers \
             that merely look array-ish, so array slots reach only via the \
             runtime `ObjectKind::Array` path until step 4. Measured: PHP's \
             `array $xs` yields the hint `array`, which no shared classifier \
             reads."
        }
        T::Map => {
            "Resolves too narrowly: `is_dictionary_type_hint` requires the \
             literal substring `dictionary`, so it catches .NET but not Dart's \
             `Map<String,int>` nor a Python dict."
        }
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The default for string length must be the UTF-16 one. `common:str_length`
    /// is the byte count, and picking it here would silently reintroduce the
    /// Dart `'café'.length == 5` bug for every language at once.
    #[test]
    fn string_len_default_is_utf16_not_bytes() {
        let d = platform_defaults();
        assert_eq!(
            d.get(BuiltinType::String, ProtocolSlot::Len),
            Some("host:ecma:string:length")
        );
        assert_ne!(
            d.get(BuiltinType::String, ProtocolSlot::Len),
            Some("common:str_length")
        );
    }

    /// Every default must name a target the dispatcher can actually resolve —
    /// one of the profile's documented emit-target prefixes.
    #[test]
    fn every_default_is_a_well_formed_emit_target() {
        for (ty, slot, target) in platform_defaults().iter() {
            assert!(
                target.starts_with("common:")
                    || target.starts_with("host:")
                    || target.starts_with("opcode:")
                    || target.starts_with("intrinsic:")
                    || target.starts_with("stdlib:"),
                "{ty:?}/{slot:?} -> {target:?} is not an emit target"
            );
        }
    }

    /// A `common:` default must name something a dispatcher CLAIMS.
    ///
    /// The prefix check above is shape-only, and a well-formed name nobody
    /// dispatches emits NOTHING — stranding operands on the stack far from the
    /// table that named it. `common:dyn_ne` was added as the `Ne` default and
    /// had no arm; this is what caught it.
    ///
    /// Runs the dispatcher against a scratch chunk purely to see whether it
    /// claims the name. The emitted bytecode is discarded.
    #[test]
    fn every_common_default_is_actually_dispatched() {
        for (ty, slot, target) in platform_defaults().iter() {
            let Some(name) = target.strip_prefix("common:") else {
                continue;
            };
            let mut chunks = vec![vybe_runtime::Chunk::new("<probe>")];
            let claimed = crate::primitives::dispatch::emit_common(name, &mut chunks, 0, 2, 0);
            assert!(
                claimed,
                "{ty:?}/{slot:?} defaults to `common:{name}`, which no dispatcher \
                 claims — it would emit nothing and corrupt the stack"
            );
        }
    }

    /// A pair any LANGUAGE binds must not also claim to be unbound.
    ///
    /// `a_bound_pair_never_claims_to_be_unbound` checks only the PLATFORM
    /// table, so it cannot see a language binding — and three reasons went
    /// stale exactly that way: PHP bound `(String, Eq)`/`(String, Ne)` and
    /// Python bound `(Int, Mod)` while all three still read "no emit target
    /// exists". §5 asks for a reader OR a reason; both at once is a
    /// contradiction, and this is what enforces it.
    #[test]
    fn no_language_binding_contradicts_an_unbound_reason() {
        for lang in crate::languages::all() {
            let Ok(profile) = vybe_runtime::profile::parse_profile((lang.profile_source)()) else {
                continue;
            };
            for (ty, slot, target) in profile.builtin_slots.iter() {
                assert!(
                    unbound_reason(ty, slot).is_none(),
                    "{} binds {ty:?}/{slot:?} to {target:?}, yet `unbound_reason` \
                     still says the pair has no target — one of the two is stale",
                    lang.name
                );
            }
        }
    }

    /// `unbound_reason` means *no target exists*, so it must never explain a
    /// pair that IS in the table.
    ///
    /// Without this, `array`/`Len` could sit in the default table while also
    /// reporting a reason for having no binding, and step 4's audit of "which
    /// pairs still need bindings" would read it as handled — when what it
    /// actually needs is a resolver (`unresolvable_reason`), a different fix by
    /// a different person.
    #[test]
    fn a_bound_pair_never_claims_to_be_unbound() {
        let d = platform_defaults();
        for ty in BuiltinType::ALL {
            for (slot, target) in d.slots_for(ty) {
                assert!(
                    unbound_reason(ty, slot).is_none(),
                    "{ty:?}/{slot:?} is bound to {target:?} yet carries an \
                     unbound reason"
                );
            }
        }
    }

    /// PINS the current unresolvable set. It does NOT track `builtin_type_of`:
    /// both sides are constants, so when step 4 makes `Map` resolve through a
    /// profile-declared spelling this test keeps passing while the record goes
    /// stale. Deriving it would need a `Compiler` instance, which a unit test
    /// here has no cheap way to build.
    ///
    /// So its value is narrow and worth stating plainly: it catches a reason
    /// being *added or deleted* without thought, not a resolver that quietly
    /// grew. A step-4 reader must re-check the set by hand, not trust this.
    ///
    /// `String` and `Map` are the two types `builtin_type_of` attempts today.
    /// `Map` appears in both directions on purpose: it resolves, but only for
    /// the `dictionary` spelling, so it is unresolvable-in-general.
    #[test]
    fn unresolvable_types_are_exactly_the_ones_the_resolver_misses() {
        for ty in BuiltinType::ALL {
            let recorded = unresolvable_reason(ty).is_some();
            let expected = matches!(
                ty,
                BuiltinType::Int
                    | BuiltinType::Double
                    | BuiltinType::BigInt
                    | BuiltinType::Array
                    | BuiltinType::Map
            );
            assert_eq!(
                recorded, expected,
                "{ty:?}: unresolvable_reason recorded={recorded} expected={expected} \
                 — `builtin_type_of` changed without updating the record"
            );
        }
    }

    /// §2d precedence: a language's `[builtin_slots.*]` entry beats the
    /// platform default, and a language that declares nothing is unaffected.
    ///
    /// This is the property the whole override mechanism rests on, and it is
    /// worth pinning directly rather than only through a language's slices:
    /// `Map`/`GetItem` is bound centrally to the ECMA answer (`undefined` on a
    /// miss) ONLY because Dart can override it with its own `null`-returning
    /// emitter. If `get_or` ever resolved default-first, that binding would
    /// silently reintroduce the exact bug that kept the pair unbound.
    #[test]
    fn a_language_override_beats_the_platform_default() {
        use vybe_ast::builtin_slots::BuiltinSlotBindings;

        let mut lang = BuiltinSlotBindings::new();
        lang.insert(
            BuiltinType::Map,
            ProtocolSlot::GetItem,
            "common:dart.index_get",
        );

        // Declared → the language wins.
        assert_eq!(
            lang.get_or(defaults(), BuiltinType::Map, ProtocolSlot::GetItem),
            Some("common:dart.index_get")
        );
        // Not declared → the platform default still answers.
        assert_eq!(
            lang.get_or(defaults(), BuiltinType::Map, ProtocolSlot::Len),
            Some("common:dict.size")
        );
        // A language with an EMPTY table gets the default for everything, so a
        // language that declares nothing cannot be affected by this mechanism.
        let empty = BuiltinSlotBindings::new();
        assert_eq!(
            empty.get_or(defaults(), BuiltinType::Map, ProtocolSlot::GetItem),
            Some("common:dict.get_dynamic")
        );
    }

    /// An unbound pair is either explained or absent-by-omission. This test
    /// pins the ones we have deliberately left unbound so that removing a
    /// reason without adding a binding is caught.
    #[test]
    fn deliberately_unbound_pairs_carry_a_reason() {
        let d = platform_defaults();
        for (ty, slot) in [
            (BuiltinType::String, ProtocolSlot::Bool),
            (BuiltinType::String, ProtocolSlot::Hash),
            (BuiltinType::String, ProtocolSlot::Iterator),
        ] {
            assert!(d.get(ty, slot).is_none(), "{ty:?}/{slot:?} is bound");
            assert!(
                unbound_reason(ty, slot).is_some(),
                "{ty:?}/{slot:?} is unbound with no reason"
            );
        }

        // The other direction, for the pairs that GRADUATED from unbound to
        // bound once per-language overrides existed. `a_bound_pair_never_claims_
        // to_be_unbound` only catches re-adding a reason while still bound; this
        // catches the reverse — unbinding one of these without noticing that a
        // language's override is what makes the central answer safe.
        for (ty, slot) in [
            (BuiltinType::Map, ProtocolSlot::GetItem),
            (BuiltinType::Array, ProtocolSlot::GetItem),
            (BuiltinType::String, ProtocolSlot::GetItem),
            (BuiltinType::Bytes, ProtocolSlot::GetItem),
            (BuiltinType::String, ProtocolSlot::Eq),
            (BuiltinType::String, ProtocolSlot::Ne),
        ] {
            assert!(
                d.get(ty, slot).is_some(),
                "{ty:?}/{slot:?} became unbound — if that is intended it needs \
                 an `unbound_reason`, and any language override of it is now dead"
            );
        }
    }
}

/// The default table, built once. `platform_defaults()` allocates, and slot
/// resolution happens per expression.
fn defaults() -> &'static BuiltinSlotBindings {
    static DEFAULTS: OnceLock<BuiltinSlotBindings> = OnceLock::new();
    DEFAULTS.get_or_init(platform_defaults)
}

impl Compiler {
    /// The built-in type of `expr`, when the compiler can name it statically.
    ///
    /// This is the COMPILE-TIME half of `builtinslotplan.md` §2c. `None` means
    /// either a user type or one the platform cannot currently classify; both
    /// fall to the runtime path, which is unchanged.
    ///
    /// # Why this delegates instead of matching
    ///
    /// The platform ALREADY holds the per-language spelling knowledge, in
    /// `Compiler::is_string_type_hint` (`string`, `system.string`, `*.string`,
    /// `character`, `character(N)`, `character*N`) and
    /// `Compiler::is_dictionary_type_hint`. Writing a second list here —
    /// `"String" | "str" | "System.String" | ...` — would put a per-language
    /// table in shared code AND let the two lists drift, so the census would
    /// measure this function's opinion rather than the platform's behaviour.
    ///
    /// Those two predicates (plus `is_numeric_type_hint`) are precisely what
    /// step 4's profile section is chartered to replace; delegating keeps the
    /// eventual move a single-site change.
    /// Step 4 made this a single table lookup. It consults the profile's
    /// `[builtin_types]` spellings first, then the platform table — the same
    /// language-first precedence as `BuiltinSlotBindings::get_or`.
    ///
    /// The numeric and array cases that step 3 recorded as unresolvable resolve
    /// here now: `classify_with` returns WHICH built-in a hint names, where the
    /// old `is_numeric_type_hint` could only say "some number".
    pub(crate) fn builtin_type_of(&self, expr: &Expression) -> Option<BuiltinType> {
        let hint = self.infer_expr_type_hint(expr)?;
        vybe_ast::builtin_types::classify_with(&self.profile.builtin_type_spellings, &hint)
    }

    /// The emit target bound for `(static type of expr, slot)`, or `None` when
    /// the type is not a built-in or the pair is unbound.
    ///
    /// Language table first, platform default second — §2d steps 2 and 3.
    #[allow(dead_code)]
    pub(crate) fn builtin_slot_target(
        &self,
        expr: &Expression,
        slot: ProtocolSlot,
    ) -> Option<&str> {
        let ty = self.builtin_type_of(expr)?;
        self.profile.builtin_slots.get_or(defaults(), ty, slot)
    }
}

impl Compiler {
    /// Rewrite a matched value-method's emit target to the slot binding for its
    /// receiver's built-in type — `builtinslotplan.md` step 5.
    ///
    /// This is the point where the table finally *decides* something. Everything
    /// before it (steps 1–4) established that the table faithfully describes
    /// current behaviour; this substitutes the table's answer for the profile's.
    ///
    /// Three conditions must all hold, and each is a deliberate gate:
    ///
    /// 1. the profile declared `slot = "..."` on the method — the LANGUAGE says
    ///    which slot its spelling fills, so no method-name table exists here;
    /// 2. the receiver's built-in type is statically known (§2c compile-time
    ///    path);
    /// 3. the `(type, slot)` pair is bound.
    ///
    /// Any one failing leaves `def` exactly as the profile wrote it, which is
    /// why a language that declares nothing cannot be affected.
    ///
    /// Returns the def unchanged when the substitution does not apply.
    pub(crate) fn apply_builtin_slot_binding(
        &self,
        object: &Expression,
        mut def: vybe_runtime::profile::BuiltinDef,
    ) -> vybe_runtime::profile::BuiltinDef {
        let Some(slot) = def.slot else { return def };
        let Some(ty) = self.builtin_type_of(object) else {
            return def;
        };
        // §2d precedence: the language's own `[builtin_slots.*]` entry wins over
        // the platform default. This is what makes a slot bindable at all where
        // languages genuinely disagree — `Map`/`GetItem`'s four different
        // answers on a miss, `Eq`'s per-language structural rules.
        let Some(target) = self.profile.builtin_slots.get_or(defaults(), ty, slot) else {
            return def;
        };
        // Reuses the profile's own emit-target parser rather than a second
        // interpreter: a slot binding is deliberately not a new vocabulary.
        if let Some(emit) = vybe_runtime::profile::parse_emit_target(target) {
            def.emit = emit;
        }
        def
    }

    /// Record `(language, built-in receiver type, method, emit target)` for a
    /// value-method call whose receiver has a statically-known built-in type.
    ///
    /// Deliberately a census and not a comparison: mapping a method NAME to a
    /// slot would need a name table in shared code, which is the very
    /// anti-pattern this plan exists to remove. The output is read by a human
    /// (or a script) to decide which pairs deserve bindings, and the language
    /// keeps owning its spellings.
    pub(crate) fn audit_builtin_slot_census(
        &self,
        object: &Expression,
        method: &str,
        emit: &impl std::fmt::Debug,
    ) {
        // Read once. This runs on EVERY builtin call and EVERY value-method
        // dispatch, and `env::var` allocates a String each time — §5 requires
        // the static path to cost less than today's probe, so the audit must
        // not quietly add per-call allocation to the build it is measuring.
        static ENABLED: OnceLock<bool> = OnceLock::new();
        if !ENABLED.get_or_init(|| std::env::var_os("VYBE_SLOT_AUDIT").is_some()) {
            return;
        }
        // An unresolved receiver is recorded too, and is the more important
        // number: it says how much of the surface the COMPILE-TIME path (§2c)
        // can see at all. A binding only ever applies where the type is known,
        // so `?` rows are the measured ceiling on step 5's reach.
        let hint = self.infer_expr_type_hint(object);
        let ty = self
            .builtin_type_of(object)
            .map(|t| t.as_key())
            .unwrap_or("?");
        eprintln!(
            "[slot-census] lang={} type={} hint={} method={} emit={:?}",
            self.profile.name,
            ty,
            hint.as_deref().unwrap_or("-"),
            method,
            emit
        );
    }
}

#[cfg(test)]
mod resolver_tests {
    use super::*;

    #[test]
    fn defaults_are_built_once() {
        assert!(std::ptr::eq(defaults(), defaults()));
    }

    /// A literal's type is known, so a bound slot resolves without any runtime
    /// probe — this is the whole point of the compile-time path.
    #[test]
    fn a_string_literal_resolves_its_bound_slots() {
        let d = defaults();
        assert_eq!(
            d.get(BuiltinType::String, ProtocolSlot::Len),
            Some("host:ecma:string:length")
        );
        assert_eq!(
            d.get(BuiltinType::String, ProtocolSlot::GetItem),
            Some("common:str_char_at")
        );
    }

    /// `from_type_hint` is the gate: a user type must never resolve to a
    /// built-in binding, or a class named `String` would inherit string
    /// semantics it never asked for (§2d).
    #[test]
    fn user_types_do_not_resolve_to_builtins() {
        assert_eq!(BuiltinType::from_type_hint("Toggle"), None);
        assert_eq!(BuiltinType::from_type_hint("MyString"), None);
    }
}
