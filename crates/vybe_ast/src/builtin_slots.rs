//! Protocol-slot bindings for BUILT-IN types.
//!
//! `flexclassplan.md` §2b gives three ways to register a slot binding, and all
//! three assume a **class**. A `string`, an `int`, an array or a map cannot
//! declare anything — so wherever a built-in's behaviour differs per language,
//! the platform forked instead of binding:
//!
//! - `primitives/strings.rs::emit_length` hardcodes `wasm:js-string.length`
//!   (UTF-8 bytes), so Dart's `'café'.length` was 5 until Dart was routed
//!   *away* from the shared helper.
//! - `LanguageHooks::value_eq` and `::relational_compare` are private callbacks
//!   for the two languages that noticed their comparison differed.
//!
//! This module is the vocabulary for saying it instead: a binding is
//! `(BuiltinType, ProtocolSlot) -> emit target`, where the emit target is the
//! SAME string vocabulary the profile already uses for `[builtins]` and
//! `[value_methods]` (`opcode:` / `intrinsic:` / `common:` / `host:` /
//! `stdlib:`). No new vocabulary — a new table, not a new language.
//!
//! Nothing reads this yet. It is step 1 of `builtinslotplan.md`, which is
//! deliberately inert: steps 1–4 establish that the table faithfully describes
//! current behaviour BEFORE any of it changes.

use crate::ProtocolSlot;
use std::collections::HashMap;

/// A built-in type that can carry slot bindings.
///
/// The variants are exactly the types the platform can already name, from both
/// directions — nothing here is invented:
///
/// - **statically**, by `type_inference::infer_expr_type_hint`, which returns
///   `"string"`, `"int"`, `"integer"`, `"double"`, `"bigint"`, `"bool"`,
///   `"char"`;
/// - **at runtime**, by `vybe_runtime::value::ObjectKind`, which has exactly
///   `Ordinary`, `Array`, `Map`, `Set`, `ArrayBuffer`.
///
/// Those are the two resolution paths (`builtinslotplan.md` §2c), and a type
/// is listed here only if at least one of them can identify it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BuiltinType {
    String,
    Int,
    Double,
    BigInt,
    Bool,
    Char,
    Array,
    Map,
    Set,
    /// Binary data — `ObjectKind::ArrayBuffer`. Distinct from `String`: bytes
    /// are a different TYPE, not a differently-encoded string.
    Bytes,
    /// A plain object — `ObjectKind::Ordinary`. Bound last, because a user
    /// class always wins over it (§2d).
    Object,
}

impl BuiltinType {
    /// The profile section key: `[builtin_slots.string]`.
    pub fn as_key(self) -> &'static str {
        match self {
            BuiltinType::String => "string",
            BuiltinType::Int => "int",
            BuiltinType::Double => "double",
            BuiltinType::BigInt => "bigint",
            BuiltinType::Bool => "bool",
            BuiltinType::Char => "char",
            BuiltinType::Array => "array",
            BuiltinType::Map => "map",
            BuiltinType::Set => "set",
            BuiltinType::Bytes => "bytes",
            BuiltinType::Object => "object",
        }
    }

    pub fn from_key(key: &str) -> Option<Self> {
        Some(match key {
            "string" => BuiltinType::String,
            "int" => BuiltinType::Int,
            "double" => BuiltinType::Double,
            "bigint" => BuiltinType::BigInt,
            "bool" => BuiltinType::Bool,
            "char" => BuiltinType::Char,
            "array" => BuiltinType::Array,
            "map" => BuiltinType::Map,
            "set" => BuiltinType::Set,
            "bytes" => BuiltinType::Bytes,
            "object" => BuiltinType::Object,
            _ => return None,
        })
    }

    /// Resolve the COMPILE-TIME path: a static type hint to the built-in it
    /// names, or `None` when the hint names a user type (or nothing).
    ///
    /// `"integer"` and `"int"` both map to [`BuiltinType::Int`] because
    /// `infer_expr_type_hint` produces both spellings.
    pub fn from_type_hint(hint: &str) -> Option<Self> {
        Some(match hint.trim() {
            "string" | "str" => BuiltinType::String,
            "int" | "integer" => BuiltinType::Int,
            "double" | "float" => BuiltinType::Double,
            "bigint" => BuiltinType::BigInt,
            "bool" | "boolean" => BuiltinType::Bool,
            "char" => BuiltinType::Char,
            _ => return None,
        })
    }

    /// Every built-in, for exhaustive iteration when building or auditing the
    /// default table.
    pub const ALL: [BuiltinType; 11] = [
        BuiltinType::String,
        BuiltinType::Int,
        BuiltinType::Double,
        BuiltinType::BigInt,
        BuiltinType::Bool,
        BuiltinType::Char,
        BuiltinType::Array,
        BuiltinType::Map,
        BuiltinType::Set,
        BuiltinType::Bytes,
        BuiltinType::Object,
    ];
}

/// How a `(BuiltinType, ProtocolSlot)` pair is emitted.
///
/// This is deliberately the profile's existing emit-target string rather than a
/// parsed enum: `opcode:f64_abs`, `common:strings.len_bytes`,
/// `host:ecma:string:length`. The dispatcher that resolves those strings
/// already exists and already documents its precedence, so a binding needs no
/// interpretation of its own.
pub type EmitTarget = String;

/// The `(BuiltinType, ProtocolSlot) -> EmitTarget` table.
///
/// One of these is the platform DEFAULT table (`builtinslotplan.md` §2e); each
/// language may carry a second, sparse one that overrides it. Measured
/// 2026-07-30: for string slots, PHP is the only language that needs any
/// entries at all — every other language agrees with ECMA on `Len`, `Eq`,
/// `Bool`, `Lt`, `Contains` and `Add`.
#[derive(Debug, Clone, Default)]
pub struct BuiltinSlotBindings {
    bindings: HashMap<(BuiltinType, ProtocolSlot), EmitTarget>,
}

impl BuiltinSlotBindings {
    pub fn new() -> Self {
        Self::default()
    }

    /// Record a binding, returning the target it displaced.
    ///
    /// A duplicate is not rejected here: a language's table legitimately
    /// overrides the default table for the same pair. Conflict *within* one
    /// table is a profile error and is reported by the profile reader, which
    /// is the layer that knows the source location.
    pub fn insert(
        &mut self,
        ty: BuiltinType,
        slot: ProtocolSlot,
        target: impl Into<EmitTarget>,
    ) -> Option<EmitTarget> {
        self.bindings.insert((ty, slot), target.into())
    }

    pub fn get(&self, ty: BuiltinType, slot: ProtocolSlot) -> Option<&str> {
        self.bindings.get(&(ty, slot)).map(String::as_str)
    }

    /// This table's answer, else `fallback`'s. The order is the language's
    /// table first, then the platform default — §2d steps 2 and 3.
    pub fn get_or<'a>(
        &'a self,
        fallback: &'a BuiltinSlotBindings,
        ty: BuiltinType,
        slot: ProtocolSlot,
    ) -> Option<&'a str> {
        self.get(ty, slot).or_else(|| fallback.get(ty, slot))
    }

    pub fn is_empty(&self) -> bool {
        self.bindings.is_empty()
    }

    pub fn len(&self) -> usize {
        self.bindings.len()
    }

    pub fn iter(&self) -> impl Iterator<Item = (BuiltinType, ProtocolSlot, &str)> {
        self.bindings
            .iter()
            .map(|((ty, slot), target)| (*ty, *slot, target.as_str()))
    }

    /// Slots bound for `ty`, for auditing a single row of the table.
    pub fn slots_for(&self, ty: BuiltinType) -> impl Iterator<Item = (ProtocolSlot, &str)> {
        self.bindings
            .iter()
            .filter(move |((t, _), _)| *t == ty)
            .map(|((_, slot), target)| (*slot, target.as_str()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn type_hint_spellings_both_map_to_int() {
        assert_eq!(BuiltinType::from_type_hint("int"), Some(BuiltinType::Int));
        assert_eq!(
            BuiltinType::from_type_hint("integer"),
            Some(BuiltinType::Int)
        );
    }

    #[test]
    fn a_user_type_name_is_not_a_builtin() {
        assert_eq!(BuiltinType::from_type_hint("Toggle"), None);
    }

    #[test]
    fn key_roundtrips_for_every_builtin() {
        for ty in BuiltinType::ALL {
            assert_eq!(BuiltinType::from_key(ty.as_key()), Some(ty));
        }
    }

    #[test]
    fn language_table_overrides_default_then_falls_back() {
        let mut default = BuiltinSlotBindings::new();
        default.insert(BuiltinType::String, ProtocolSlot::Len, "host:ecma:string:length");
        default.insert(BuiltinType::String, ProtocolSlot::Eq, "common:ops.eq_exact");

        // PHP is the one language the 2026-07-30 matrix showed diverging.
        let mut php = BuiltinSlotBindings::new();
        php.insert(BuiltinType::String, ProtocolSlot::Len, "common:strings.len_bytes");

        assert_eq!(
            php.get_or(&default, BuiltinType::String, ProtocolSlot::Len),
            Some("common:strings.len_bytes")
        );
        assert_eq!(
            php.get_or(&default, BuiltinType::String, ProtocolSlot::Eq),
            Some("common:ops.eq_exact")
        );
        assert_eq!(
            php.get_or(&default, BuiltinType::Map, ProtocolSlot::Len),
            None
        );
    }
}
