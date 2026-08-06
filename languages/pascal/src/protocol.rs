//! Pascal source spelling -> shared protocol slot.
//!
//! Pascal-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is Pascal's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a Pascal member name to `(canonical, slot?)`.
///
/// Pascal resolves members case-insensitively, so everything is matched
/// lowercased. `class operator Add` arrives as `"Add"` from the
/// operator-overload grammar rule; `destructor Destroy` as `"Destroy"`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name.to_lowercase().as_str() {
        "destroy" => ("destructor".into(), Some(Destructor)),
        "tostring" => ("tostring".into(), Some(ToString)),

        // ── Roles Pascal spells, that the shared compiler already resolves ──
        //
        // These were the gap. Pascal declared only `Destructor` and
        // `ToString`, while every other language declares `Iterator`, and
        // most declare `Compare`/`Hash`/`GetItem`/`Len`. The consequence was
        // not that the features were missing — it was that Pascal
        // re-implemented each one locally, so a Pascal class could not answer
        // an iteration or comparison request from anywhere else.
        //
        // Delphi shares C#'s vocabulary almost exactly, so the canonical
        // names here match `languages/csharp/src/protocol.rs`.
        // ── ONLY CONTRACT SPELLINGS BELOW ──────────────────────────────
        //
        // `GetEnumerator`/`MoveNext` (the `for..in` contract), `CompareTo`
        // (IComparable) and `GetHashCode` are RESERVED by contract in Delphi:
        // a type that spells one is declaring the role.
        //
        // `Contains`, `GetItem`, `SetItem`, `GetCount`, `Equals` and `Assign`
        // are NOT — they are ordinary method names. I mapped them anyway and
        // it cost **369 tests**: every class spelling `Contains` claimed the
        // Contains slot, breaking `x in someSet`; `Equals` claimed `Eq`,
        // breaking `class operator` dispatch on records. This is
        // flexclassplan §1e measured — "any class defining `__str__` silently
        // claims all three".
        //
        // Those roles are real and Pascal does fill them, but from the
        // DECLARATION, not the spelling: the `default` directive on an indexed
        // property is GetItem/SetItem, `Count` on a declared collection is
        // Len. That needs `Modifiers.protocol_slot` (currently zero producers
        // and zero consumers) — a walker STATING the role. Do not re-add a
        // name row here as a shortcut; it has been measured and it is worse
        // than the gap.
        "getenumerator" => ("iterator".into(), Some(Iterator)),
        "movenext" => ("next".into(), Some(Next)),
        "compareto" => ("compare".into(), Some(Compare)),
        "gethashcode" => ("hash".into(), Some(Hash)),
        // `Equals` is deliberately NOT mapped to `Eq`. I mapped it and it
        // broke every record: `TObject.Equals` is REFERENCE equality, while
        // `class operator Equal` is the `=` overload — two different
        // operations. Since every Pascal class inherits `Equals`, mapping it
        // made every class claim the `Eq` slot, so the operator dispatch found
        // `TObject.Equals` instead of the declared operator. Measured:
        // `class operator BitwiseAnd` returned `[object Object]` where fpc
        // gives `10`. This is flexclassplan §1e exactly — "any class defining
        // `__str__` silently claims all three".
        // A STOPGAP, and known to be the wrong mechanism. Delphi marks the
        // subscript role with the `default` DIRECTIVE on an indexed property
        // — `property Items[i: Integer]: T read GetElement; default;` — and
        // the accessor may be named anything. Matching the conventional
        // spelling is the name-as-identity pattern flexclassplan §1e exists
        // to remove.
        //
        // The principled route is `Modifiers.protocol_slot`, which §"STATUS
        // 2026-07-28" added so "a walker can now STATE the role instead of
        // leaving a name to be re-matched". **That field has zero producers
        // and zero consumers repo-wide** — it was declared and never wired at
        // either end. Until it is, this is the only route that reaches
        // `expressions.rs`/`statements.rs`, whose subscript dispatch already
        // substitutes the slot. Pascal meanwhile carries the fact as a magic
        // decorator string (`"__pascal_default_property"`) feeding a
        // `PascalIndexedPropertyInfo` side table — the very "name-matched side
        // table" §2c-bis says to replace.
        // `Count` is Delphi's length spelling everywhere — `TList`,
        // `TStringList`, `TCollection`. It already fills the Len slot for the
        // BUILT-IN types via `[value_methods] slot = "len"`; this is the same
        // role for a USER class, which that table cannot reach.
        // `Assign` is Delphi's copy protocol (`TPersistent.Assign`).
        "add" => ("add".into(), Some(Add)),
        "subtract" => ("sub".into(), Some(Sub)),
        "multiply" => ("mul".into(), Some(Mul)),
        "divide" => ("div".into(), Some(Div)),
        "intdivide" => ("floordiv".into(), Some(FloorDiv)),
        "modulus" => ("mod".into(), Some(Mod)),
        "negative" => ("neg".into(), Some(Neg)),
        "positive" => ("pos".into(), Some(Pos)),
        "equal" => ("eq".into(), Some(Eq)),
        "notequal" => ("ne".into(), Some(Ne)),
        "lessthan" => ("lt".into(), Some(Lt)),
        "lessthanorequal" => ("le".into(), Some(Le)),
        "greaterthan" => ("gt".into(), Some(Gt)),
        "greaterthanorequal" => ("ge".into(), Some(Ge)),
        "bitwiseand" => ("and".into(), Some(And)),
        "bitwiseor" => ("or".into(), Some(Or)),
        "bitwisexor" => ("xor".into(), Some(Xor)),
        "logicalnot" | "bitwisenot" => ("not".into(), Some(Not)),
        "leftshift" => ("lshift".into(), Some(LShift)),
        "rightshift" => ("rshift".into(), Some(RShift)),
        other => (other.to_string(), None),
    }
}
