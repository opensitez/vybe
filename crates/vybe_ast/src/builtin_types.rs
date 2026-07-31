//! Which SOURCE TYPE NAMES name a built-in — `builtinslotplan.md` step 4.
//!
//! `builtin_slots.rs` says what a built-in type does for a slot. This says how
//! to recognise one in the first place, which step 3's census measured to be
//! the binding constraint: a binding only ever applies where the receiver's
//! type can be named, and today most receivers cannot be.
//!
//! # What this replaces
//!
//! Three predicates in `vybe_compiler`, each a per-language spelling table
//! living in shared code:
//!
//! | function | held |
//! |---|---|
//! | `Compiler::is_string_type_hint` | `string`, `system.string`, `*.string`, `character`, `character(N)`, `character*N` |
//! | `Compiler::is_numeric_type_hint` | ~20 spellings across .NET / Pascal / Fortran |
//! | `Compiler::is_dictionary_type_hint` | `dictionary`, `hashtable` |
//!
//! They stay as thin delegators so their call sites do not move, but the
//! knowledge lives here, in a table a profile can extend.
//!
//! # Why the defaults keep every spelling for every language
//!
//! Because that is what the platform does today, and step 4 must be neutral by
//! construction. `is_string_type_hint` answers `true` for `character(20)` in
//! JavaScript right now — nonsense, but harmless nonsense, since no JS program
//! produces that hint. Narrowing it per language is a step 5 decision that
//! wants its own measurement; doing it here would smuggle a behaviour change
//! into a move.
//!
//! What step 4 DOES fix is the other direction — spellings that no shared list
//! contains, so an annotated program resolves nothing. Step 3 measured two:
//! Python's `str` and PHP's `array`. Those arrive via a profile's
//! `[builtin_types]` section, not by being appended here.

use crate::builtin_slots::BuiltinType;
use std::borrow::Cow;

/// How a spelling is matched. The three shapes are exactly those the moved
/// predicates used — nothing is invented, and nothing is generalised into a
/// pattern language.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Match {
    /// The whole hint equals the pattern (`int`, `system.string`).
    Exact,
    /// The hint ends with the pattern (`.string`, catching a namespaced
    /// `My.Ns.String`).
    Suffix,
    /// The hint starts with the pattern (`character(`, `character*` — Fortran
    /// carries the length inside the type name).
    Prefix,
    /// The pattern occurs anywhere (`dictionary`, catching
    /// `System.Collections.Generic.Dictionary<K,V>`).
    Contains,
}

/// One spelling and how to match it.
///
/// `pattern` is a [`Cow`] so the platform table stays a `const` of borrowed
/// literals while a profile can contribute owned strings parsed from TOML —
/// one type for both sources, so `classify_with` chains them directly.
#[derive(Debug, Clone)]
pub struct Spelling {
    pub pattern: Cow<'static, str>,
    pub how: Match,
    pub ty: BuiltinType,
}

/// A borrowed spelling, for the `const` platform table.
const fn s(pattern: &'static str, how: Match, ty: BuiltinType) -> Spelling {
    Spelling {
        pattern: Cow::Borrowed(pattern),
        how,
        ty,
    }
}

impl Spelling {
    /// An owned spelling, for a profile's `[builtin_types]` section.
    ///
    /// The pattern is normalised on the way in, so a profile writing `String`
    /// or ` Str ` matches the same hints as one writing `string` — the table is
    /// always compared against a normalised hint.
    pub fn owned(pattern: &str, how: Match, ty: BuiltinType) -> Self {
        Spelling {
            pattern: Cow::Owned(normalize(pattern)),
            how,
            ty,
        }
    }
}

/// The platform's spellings, moved verbatim from the three predicates.
///
/// ORDER MATTERS: the first match wins, and the numeric entries are ordered so
/// that a language's own spelling is not shadowed by a substring of another's.
/// All matching is done on a lowercased, trimmed hint — see [`normalize`].
pub const PLATFORM_SPELLINGS: &[Spelling] = &[
    // ── string: from `Compiler::is_string_type_hint` ─────────────────────
    s("string", Match::Exact, BuiltinType::String),
    s("system.string", Match::Exact, BuiltinType::String),
    s(".string", Match::Suffix, BuiltinType::String),
    s("character", Match::Exact, BuiltinType::String),
    s("character(", Match::Prefix, BuiltinType::String),
    s("character*", Match::Prefix, BuiltinType::String),
    // ── bytes ───────────────────────────────────────────────────────────
    // `Literal::Bytes` gives these a static type to infer (unifiedstringplan
    // §3c). Bytes are a distinct TYPE, not an encoding of `String`, so they
    // classify separately rather than as a string spelling.
    s("bytes", Match::Exact, BuiltinType::Bytes),
    s("bytearray", Match::Exact, BuiltinType::Bytes),
    s("uint8array", Match::Exact, BuiltinType::Bytes),
    // ── char ────────────────────────────────────────────────────────────
    // `is_numeric_type_hint` does NOT list `char`, and neither did the old
    // string predicate, so it belongs to neither. It is named here because
    // `BuiltinType::Char` exists and a hint of `char` should resolve to it
    // rather than to nothing.
    s("char", Match::Exact, BuiltinType::Char),
    // ── bool ────────────────────────────────────────────────────────────
    // Spelled out in `emit_default_value_for_type_hint`'s match rather than in
    // a predicate of its own; same two spellings.
    s("bool", Match::Exact, BuiltinType::Bool),
    s("boolean", Match::Exact, BuiltinType::Bool),
    // ── numerics: from `Compiler::is_numeric_type_hint` ──────────────────
    //
    // That predicate returned a single bool for all twenty spellings, which is
    // why step 3 recorded `int`/`double`/`bigint` as UNRESOLVABLE. Splitting
    // the list is the whole point: the spellings are unchanged, but each now
    // says which numeric it names.
    //
    // The int/double split follows the language's own type, not the WASM
    // representation — every numeric is an f64 at runtime. `real`, `double`,
    // `float`, `single` and `decimal` are the languages' floating types;
    // the rest are integral.
    s("int", Match::Exact, BuiltinType::Int),
    s("integer", Match::Exact, BuiltinType::Int),
    s("int32", Match::Exact, BuiltinType::Int),
    s("longint", Match::Exact, BuiltinType::Int),
    s("long", Match::Exact, BuiltinType::Int),
    s("int64", Match::Exact, BuiltinType::Int),
    s("short", Match::Exact, BuiltinType::Int),
    s("int16", Match::Exact, BuiltinType::Int),
    s("uint", Match::Exact, BuiltinType::Int),
    s("uint32", Match::Exact, BuiltinType::Int),
    s("ulong", Match::Exact, BuiltinType::Int),
    s("uint64", Match::Exact, BuiltinType::Int),
    s("ushort", Match::Exact, BuiltinType::Int),
    s("uint16", Match::Exact, BuiltinType::Int),
    s("byte", Match::Exact, BuiltinType::Int),
    s("sbyte", Match::Exact, BuiltinType::Int),
    s("real", Match::Exact, BuiltinType::Double),
    s("double", Match::Exact, BuiltinType::Double),
    s("float", Match::Exact, BuiltinType::Double),
    s("single", Match::Exact, BuiltinType::Double),
    s("decimal", Match::Exact, BuiltinType::Double),
    // ── map: from `Compiler::is_dictionary_type_hint` ────────────────────
    s("dictionary", Match::Contains, BuiltinType::Map),
    s("hashtable", Match::Suffix, BuiltinType::Map),
];

/// Lowercase and trim — the same normalisation
/// `Compiler::normalize_type_hint` applies, reproduced here so this table can
/// be matched without a `Compiler`.
pub fn normalize(hint: &str) -> String {
    hint.trim().to_lowercase()
}

fn matches(sp: &Spelling, normalized: &str) -> bool {
    let pattern = sp.pattern.as_ref();
    match sp.how {
        Match::Exact => normalized == pattern,
        Match::Suffix => normalized.ends_with(pattern),
        Match::Prefix => normalized.starts_with(pattern),
        Match::Contains => normalized.contains(pattern),
    }
}

/// The SINGLE built-in `hint` best names, consulting `extra` (a profile's
/// declared spellings) BEFORE the platform table.
///
/// Language-first, exactly like `BuiltinSlotBindings::get_or`: a language that
/// declares `str` gets it even though the platform table has no such entry, and
/// a language that wants `real` to mean something else can say so without
/// editing shared code.
///
/// # First match wins, and that is a new decision
///
/// The predicates this table replaced were INDEPENDENT — nothing stopped a
/// hint from satisfying two of them, and `Dictionary.String` satisfied both
/// `is_string_type_hint` (ends with `.string`) and `is_dictionary_type_hint`
/// (contains `dictionary`). A resolver cannot return two types, so this
/// function picks the first.
///
/// That tie-break is why [`matches_type`] exists separately: the predicates
/// ask "could this be a string?", the resolver asks "what IS this?", and only
/// the resolver needs to choose. Answering both with `classify` would silently
/// change the predicates for overlapping hints — the differential tests below
/// caught exactly that.
pub fn classify_with(extra: &[Spelling], hint: &str) -> Option<BuiltinType> {
    let normalized = normalize(hint);
    extra
        .iter()
        .chain(PLATFORM_SPELLINGS.iter())
        .find(|sp| matches(sp, &normalized))
        .map(|sp| sp.ty)
}

/// The built-in `hint` names according to the platform table alone.
pub fn classify(hint: &str) -> Option<BuiltinType> {
    classify_with(&[], hint)
}

/// Whether ANY spelling for `ty` matches `hint`, independent of whether some
/// other type also matches.
///
/// This is the predicates' semantics, not the resolver's — see [`classify_with`].
pub fn matches_type_with(extra: &[Spelling], hint: &str, ty: BuiltinType) -> bool {
    let normalized = normalize(hint);
    extra
        .iter()
        .chain(PLATFORM_SPELLINGS.iter())
        .any(|sp| sp.ty == ty && matches(sp, &normalized))
}

/// Whether `hint` names the given built-in, platform table only.
///
/// The shape the three moved predicates need, so each becomes a one-line
/// delegation with its original semantics intact.
pub fn is(hint: &str, ty: BuiltinType) -> bool {
    matches_type_with(&[], hint, ty)
}

/// Whether `hint` names any numeric built-in — the exact question the old
/// `is_numeric_type_hint` answered, preserved for its call sites.
pub fn is_numeric(hint: &str) -> bool {
    is(hint, BuiltinType::Int) || is(hint, BuiltinType::Double) || is(hint, BuiltinType::BigInt)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every spelling the old `is_string_type_hint` accepted must still
    /// classify as a string. This is the neutrality check for the move: the
    /// list below is transcribed from that function's body, not from this
    /// table, so a dropped entry fails here.
    #[test]
    fn the_moved_string_spellings_all_survive() {
        for hint in [
            "string",
            "String",
            "system.string",
            "System.String",
            "My.Ns.String",
            "character",
            "CHARACTER",
            "character(20)",
            "character*8",
        ] {
            assert!(
                is(hint, BuiltinType::String),
                "{hint:?} was a string before the move and is not now"
            );
        }
    }

    /// Same for the twenty numeric spellings.
    #[test]
    fn the_moved_numeric_spellings_all_survive() {
        for hint in [
            "integer", "int", "int32", "longint", "real", "double", "float", "single", "decimal",
            "long", "int64", "short", "int16", "uint", "uint32", "ulong", "uint64", "ushort",
            "uint16", "byte", "sbyte",
        ] {
            assert!(
                is_numeric(hint),
                "{hint:?} was numeric before the move and is not now"
            );
        }
    }

    /// The gain the split delivers: numerics now say WHICH numeric, which is
    /// what step 3 recorded as unresolvable.
    #[test]
    fn numerics_now_discriminate_int_from_double() {
        assert_eq!(classify("int32"), Some(BuiltinType::Int));
        assert_eq!(classify("longint"), Some(BuiltinType::Int));
        assert_eq!(classify("sbyte"), Some(BuiltinType::Int));
        assert_eq!(classify("real"), Some(BuiltinType::Double));
        assert_eq!(classify("single"), Some(BuiltinType::Double));
        assert_eq!(classify("decimal"), Some(BuiltinType::Double));
    }

    #[test]
    fn the_moved_dictionary_spellings_all_survive() {
        for hint in [
            "dictionary",
            "Dictionary<string, int>",
            "System.Collections.Generic.Dictionary<K,V>",
            "Hashtable",
        ] {
            assert!(
                is(hint, BuiltinType::Map),
                "{hint:?} was a dictionary before the move and is not now"
            );
        }
    }

    /// A user type must not become a built-in — the guard that lets a class
    /// named `Toggle` keep its own slots (§2d).
    #[test]
    fn user_types_still_classify_as_nothing() {
        assert_eq!(classify("Toggle"), None);
        assert_eq!(classify("MyString"), None);
        assert_eq!(classify(""), None);
    }

    /// A profile's spellings win, and reach types the platform table cannot
    /// name at all. These two are the exact gaps step 3 measured.
    #[test]
    fn profile_spellings_are_consulted_first() {
        let python = [s("str", Match::Exact, BuiltinType::String)];
        assert_eq!(classify("str"), None, "platform table should not know `str`");
        assert_eq!(
            classify_with(&python, "str"),
            Some(BuiltinType::String),
            "a profile-declared spelling must resolve"
        );

        let php = [s("array", Match::Exact, BuiltinType::Array)];
        assert_eq!(classify("array"), None);
        assert_eq!(classify_with(&php, "array"), Some(BuiltinType::Array));
    }

    /// A profile can override the platform's answer, not merely extend it —
    /// the `get_or` precedence, applied to spellings.
    #[test]
    fn a_profile_can_override_a_platform_spelling() {
        let odd = [s("real", Match::Exact, BuiltinType::BigInt)];
        assert_eq!(classify("real"), Some(BuiltinType::Double));
        assert_eq!(classify_with(&odd, "real"), Some(BuiltinType::BigInt));
    }

    // ── differential neutrality ─────────────────────────────────────────
    //
    // A test-count comparison cannot prove a MOVE is neutral: the suites are
    // slow, coarse, and only exercise the hints their sources happen to
    // contain. These reimplement the three predicates' original bodies
    // verbatim and assert the new table agrees on every input — including the
    // adversarial ones no test suite would produce.

    /// `Compiler::is_string_type_hint`, exactly as it read before the move.
    fn old_is_string(hint: &str) -> bool {
        let n = hint.trim().to_lowercase();
        n == "string"
            || n == "system.string"
            || n.ends_with(".string")
            || n == "character"
            || n.starts_with("character(")
            || n.starts_with("character*")
    }

    /// `Compiler::is_numeric_type_hint`, exactly as it read before the move.
    fn old_is_numeric(hint: &str) -> bool {
        matches!(
            hint.trim().to_lowercase().as_str(),
            "integer"
                | "int"
                | "int32"
                | "longint"
                | "real"
                | "double"
                | "float"
                | "single"
                | "decimal"
                | "long"
                | "int64"
                | "short"
                | "int16"
                | "uint"
                | "uint32"
                | "ulong"
                | "uint64"
                | "ushort"
                | "uint16"
                | "byte"
                | "sbyte"
        )
    }

    /// `Compiler::is_dictionary_type_hint`, exactly as it read before the move.
    fn old_is_dictionary(hint: &str) -> bool {
        let n = hint.trim().to_lowercase();
        n.contains("dictionary") || n.ends_with("hashtable")
    }

    /// Type hints that actually occur, plus the ones designed to break the
    /// move. The last group matters most: the old predicates were INDEPENDENT
    /// and a hint could satisfy two at once, whereas the table returns a
    /// single answer. Any hint where that difference is observable shows up
    /// here.
    const CORPUS: &[&str] = &[
        // real spellings, with the casing and padding the compiler sees
        "string", "String", "STRING", "  string  ", "system.string", "System.String",
        "My.Ns.String", "character", "Character", "character(20)", "character*8",
        "int", "Integer", "int32", "Int32", "longint", "real", "double", "Double",
        "float", "single", "decimal", "long", "int64", "short", "int16", "uint",
        "uint32", "ulong", "uint64", "ushort", "uint16", "byte", "sbyte",
        "dictionary", "Dictionary<string, int>", "Hashtable", "System.Collections.Hashtable",
        "System.Collections.Generic.Dictionary<K,V>",
        // near-misses that must stay unclassified
        "", " ", "str", "array", "Toggle", "MyString", "stringy", "instring",
        "integer32", "realm", "bytes", "characters", "dict", "Map<String,int>",
        "List<String>", "List<string>", "char", "bool", "boolean",
        // adversarial: satisfy more than one old predicate at once
        "Dictionary.String", "Hashtable.String", "stringdictionary",
        "dictionaryhashtable", "system.string.hashtable", "character(1).string",
    ];

    /// The table answers `is(_, String)` exactly where the old predicate said
    /// string — on every input, not just the ones a suite happens to produce.
    #[test]
    fn string_classification_is_unchanged_by_the_move() {
        for hint in CORPUS {
            assert_eq!(
                is(hint, BuiltinType::String),
                old_is_string(hint),
                "{hint:?}: string classification changed"
            );
        }
    }

    #[test]
    fn numeric_classification_is_unchanged_by_the_move() {
        for hint in CORPUS {
            assert_eq!(
                is_numeric(hint),
                old_is_numeric(hint),
                "{hint:?}: numeric classification changed"
            );
        }
    }

    #[test]
    fn dictionary_classification_is_unchanged_by_the_move() {
        for hint in CORPUS {
            assert_eq!(
                is(hint, BuiltinType::Map),
                old_is_dictionary(hint),
                "{hint:?}: dictionary classification changed"
            );
        }
    }

    /// The predicates OVERLAP and the resolver does not — the distinction that
    /// `is` and `classify` exist separately to preserve.
    ///
    /// `Dictionary.String` satisfies both original predicates. `is` must keep
    /// saying yes to both, because that is what the call sites were built on;
    /// `classify` must pick one, because a resolver cannot return two types.
    /// Collapsing `is` into `classify` broke this, and the differential tests
    /// above are what caught it.
    #[test]
    fn overlapping_hints_stay_overlapping_for_predicates_but_not_the_resolver() {
        let both = "Dictionary.String";
        assert!(is(both, BuiltinType::String));
        assert!(is(both, BuiltinType::Map));
        assert_eq!(
            classify(both),
            Some(BuiltinType::String),
            "the resolver must still commit to exactly one answer"
        );
    }
}
