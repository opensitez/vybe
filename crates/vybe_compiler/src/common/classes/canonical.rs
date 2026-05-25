//! Canonical method-name table — the single source of truth for
//! "what does this language call X?".
//!
//! Every walker consults `canonicalize_method(language, source_name)`
//! to turn a source method name into a `(canonical_name,
//! special_kind?)` pair. The canonical name becomes the runtime
//! vtable key; the source name is preserved in `NormalMethod.source_name`
//! and added to `ClassType.method_aliases` so cross-language
//! dispatch still works.
//!
//! Adding a new canonical concept:
//!   1. Add a variant to `SpecialMethodKind` in `types.rs`.
//!   2. Add one match arm per language that expresses it here.
//!   3. `emit_class` learns nothing new — it already routes any
//!      canonical method to its chunk.

use super::types::SpecialMethodKind;

/// Languages whose class semantics have a canonical-name table.
/// Additive — adding a language here does not touch any other.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ClassLang {
    Js,
    Vb,
    CSharp,
    Python,
    Ruby,
    Php,
    Dart,
    Pascal,
}

/// Resolve a source method name to `(canonical, special_kind?)`.
///
/// The `canonical` name is always lowercase and language-neutral
/// (`"tostring"`, `"add"`, `"iterator"`). `None` for `special_kind`
/// means the method is ordinary — no operator / protocol role.
///
/// When the source name already matches a canonical ordinary method
/// name (e.g. a JS `render()` method), the return is
/// `(source.to_lowercase(), None)`.
pub fn canonicalize_method(
    lang: ClassLang,
    source_name: &str,
) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    // Languages that treat method names case-insensitively (VB, partly
    // Pascal) hit the vtable by lowercased name. For case-sensitive
    // languages the distinction matters at the SOURCE name only —
    // canonical is always lowercase.
    let normalized = source_name;

    match lang {
        ClassLang::Js => match normalized {
            "toString"             => ("tostring".into(), Some(ToString)),
            "valueOf"              => ("valueof".into(),  Some(ValueOf)),
            // `[Symbol.iterator]` / `[Symbol.asyncIterator]` / etc. arrive
            // from the walker as pseudo-names `Symbol.iterator` after
            // computed-key resolution.
            "Symbol.iterator"      => ("iterator".into(),      Some(Iterator)),
            "Symbol.asyncIterator" => ("asynciterator".into(), Some(AsyncIterator)),
            "Symbol.toPrimitive"   => ("toprimitive".into(),   Some(ToPrimitive)),
            "Symbol.hasInstance"   => ("hasinstance".into(),   Some(HasInstance)),
            "Symbol.toStringTag"   => ("tostringtag".into(),   None),
            _ => (normalized.to_string(), None),
        },

        ClassLang::Vb => match normalized.to_lowercase().as_str() {
            "tostring"     => ("tostring".into(),  Some(ToString)),
            "gethashcode"  => ("hash".into(),      Some(Hash)),
            "equals"       => ("eq".into(),        Some(Eq)),
            "compareto"    => ("compare".into(),   Some(Compare)),
            "getenumerator" => ("iterator".into(), Some(Iterator)),
            other => (other.to_string(), None),
        },

        ClassLang::CSharp => match normalized {
            "ToString"      => ("tostring".into(),  Some(ToString)),
            "GetHashCode"   => ("hash".into(),      Some(Hash)),
            "Equals"        => ("eq".into(),        Some(Eq)),
            "CompareTo"     => ("compare".into(),   Some(Compare)),
            "GetEnumerator" => ("iterator".into(),  Some(Iterator)),
            // `public static T operator +(...)` arrives from the walker
            // as source name `"operator+"`; same for the others.
            "operator+"  => ("add".into(),     Some(Add)),
            "operator-"  => ("sub".into(),     Some(Sub)),
            "operator*"  => ("mul".into(),     Some(Mul)),
            "operator/"  => ("div".into(),     Some(Div)),
            "operator%"  => ("mod".into(),     Some(Mod)),
            "operator==" => ("eq".into(),      Some(Eq)),
            "operator<"  => ("lt".into(),      Some(Lt)),
            "operator<=" => ("le".into(),      Some(Le)),
            "operator>"  => ("gt".into(),      Some(Gt)),
            "operator>=" => ("ge".into(),      Some(Ge)),
            _ => (normalized.to_lowercase(), None),
        },

        ClassLang::Python => match normalized {
            "__str__"         => ("tostring".into(),    Some(ToString)),
            "__repr__"        => ("repr".into(),        Some(Repr)),
            "__int__" | "__float__" => ("valueof".into(), Some(ValueOf)),
            "__iter__"        => ("iterator".into(),    Some(Iterator)),
            "__next__"        => ("next".into(),        Some(Next)),
            "__add__"         => ("add".into(),         Some(Add)),
            "__sub__"         => ("sub".into(),         Some(Sub)),
            "__mul__"         => ("mul".into(),         Some(Mul)),
            "__truediv__"     => ("div".into(),         Some(Div)),
            "__floordiv__"    => ("floordiv".into(),    Some(Div)),
            "__mod__"         => ("mod".into(),         Some(Mod)),
            "__pow__"         => ("pow".into(),         Some(Pow)),
            "__neg__"         => ("neg".into(),         Some(Neg)),
            "__eq__"          => ("eq".into(),          Some(Eq)),
            "__lt__"          => ("lt".into(),          Some(Lt)),
            "__le__"          => ("le".into(),          Some(Le)),
            "__gt__"          => ("gt".into(),          Some(Gt)),
            "__ge__"          => ("ge".into(),          Some(Ge)),
            "__and__"         => ("and".into(),         Some(And)),
            "__or__"          => ("or".into(),          Some(Or)),
            "__xor__"         => ("xor".into(),         Some(Xor)),
            "__invert__"      => ("not".into(),         Some(Not)),
            "__lshift__"      => ("lshift".into(),      Some(LShift)),
            "__rshift__"      => ("rshift".into(),      Some(RShift)),
            "__len__"         => ("len".into(),         Some(Len)),
            "__getitem__"     => ("getitem".into(),     Some(GetItem)),
            "__setitem__"     => ("setitem".into(),     Some(SetItem)),
            "__delitem__"     => ("delitem".into(),     Some(DelItem)),
            "__contains__"    => ("contains".into(),    Some(Contains)),
            "__call__"        => ("call".into(),        Some(Call)),
            "__instancecheck__" => ("hasinstance".into(), Some(HasInstance)),
            "__getattr__" | "__getattribute__" => ("getattr".into(), Some(GetAttr)),
            "__setattr__"     => ("setattr".into(),     Some(SetAttr)),
            "__delattr__"     => ("delattr".into(),     Some(DelAttr)),
            "__enter__"       => ("enter".into(),       Some(Enter)),
            "__exit__"        => ("exit".into(),        Some(Exit)),
            "__hash__"        => ("hash".into(),        Some(Hash)),
            _ => (normalized.to_string(), None),
        },

        ClassLang::Ruby => match normalized {
            "to_s"         => ("tostring".into(),  Some(ToString)),
            "inspect"      => ("repr".into(),      Some(Repr)),
            "each"         => ("iterator".into(),  Some(Iterator)),
            "+"            => ("add".into(),       Some(Add)),
            "-"            => ("sub".into(),       Some(Sub)),
            "*"            => ("mul".into(),       Some(Mul)),
            "/"            => ("div".into(),       Some(Div)),
            "%"            => ("mod".into(),       Some(Mod)),
            "**"           => ("pow".into(),       Some(Pow)),
            "-@"           => ("neg".into(),       Some(Neg)),
            "=="           => ("eq".into(),        Some(Eq)),
            "<=>"          => ("compare".into(),   Some(Compare)),
            "<"            => ("lt".into(),        Some(Lt)),
            "<="           => ("le".into(),        Some(Le)),
            ">"            => ("gt".into(),        Some(Gt)),
            ">="           => ("ge".into(),        Some(Ge)),
            "[]"           => ("getitem".into(),   Some(GetItem)),
            "[]="          => ("setitem".into(),   Some(SetItem)),
            "include?"     => ("contains".into(),  Some(Contains)),
            "size" | "length" => ("len".into(),    Some(Len)),
            "hash"         => ("hash".into(),      Some(Hash)),
            _ => (normalized.to_string(), None),
        },

        ClassLang::Php => match normalized {
            "__toString"  => ("tostring".into(),  Some(ToString)),
            "__invoke"    => ("call".into(),      Some(Call)),
            "__get"       => ("getattr".into(),   Some(GetAttr)),
            "__set"       => ("setattr".into(),   Some(SetAttr)),
            "__isset"     => ("hasattr".into(),   None), // no direct SpecialMethodKind
            "__unset"     => ("delattr".into(),   Some(DelAttr)),
            "__call"      => ("call".into(),      Some(Call)),
            _ => (normalized.to_string(), None),
        },

        ClassLang::Dart => match normalized {
            "toString"    => ("tostring".into(),  Some(ToString)),
            "hashCode"    => ("hash".into(),      Some(Hash)),
            "call"        => ("call".into(),      Some(Call)),
            // `operator +` arrives from the walker as `"operator+"`.
            "operator+"   => ("add".into(),       Some(Add)),
            "operator-"   => ("sub".into(),       Some(Sub)),
            "operator*"   => ("mul".into(),       Some(Mul)),
            "operator/"   => ("div".into(),       Some(Div)),
            "operator%"   => ("mod".into(),       Some(Mod)),
            "operator=="  => ("eq".into(),        Some(Eq)),
            "operator<"   => ("lt".into(),        Some(Lt)),
            "operator<="  => ("le".into(),        Some(Le)),
            "operator>"   => ("gt".into(),        Some(Gt)),
            "operator>="  => ("ge".into(),        Some(Ge)),
            "operator[]"  => ("getitem".into(),   Some(GetItem)),
            "operator[]=" => ("setitem".into(),   Some(SetItem)),
            _ => (normalized.to_string(), None),
        },

        ClassLang::Pascal => match normalized.to_lowercase().as_str() {
            "tostring" => ("tostring".into(), Some(ToString)),
            // Pascal `class operator Add` arrives as `"Add"` from the
            // operator-overload grammar rule.
            "add"      => ("add".into(),      Some(Add)),
            "subtract" => ("sub".into(),      Some(Sub)),
            "multiply" => ("mul".into(),      Some(Mul)),
            "divide"   => ("div".into(),      Some(Div)),
            "equal"    => ("eq".into(),       Some(Eq)),
            other => (other.to_string(), None),
        },
    }
}

/// The **reverse** lookup: given a canonical name, what aliases should
/// be installed in `ClassType.method_aliases` so callers from any
/// language find the method?
///
/// `emit_class` calls this for every `NormalMethod` that has a
/// `SpecialMethodKind`, populating the runtime alias map. Non-special
/// methods get only their source name aliased to the canonical.
pub fn aliases_for(canonical: &str, _kind: SpecialMethodKind) -> &'static [&'static str] {
    match canonical {
        "tostring" => &["toString", "ToString", "to_s", "__str__", "__toString"],
        "repr"     => &["__repr__", "inspect"],
        "hash"     => &["GetHashCode", "hashCode", "__hash__"],
        "eq"       => &["equals", "Equals", "__eq__", "operator==", "=="],
        "compare"  => &["CompareTo", "compareTo", "__cmp__", "<=>"],
        "iterator" => &["GetEnumerator", "getEnumerator", "__iter__", "each", "Symbol.iterator"],
        "next"     => &["__next__"],
        "len"      => &["length", "size", "count", "Count", "__len__"],
        "add"      => &["__add__", "operator+", "+"],
        "sub"      => &["__sub__", "operator-", "-"],
        "mul"      => &["__mul__", "operator*", "*"],
        "div"      => &["__truediv__", "operator/", "/"],
        "mod"      => &["__mod__", "%"],
        "pow"      => &["__pow__", "**"],
        "getitem"  => &["__getitem__", "operator[]", "[]"],
        "setitem"  => &["__setitem__", "operator[]=", "[]="],
        "call"     => &["__call__", "__invoke", "Invoke"],
        "contains" => &["__contains__", "include?"],
        _ => &[],
    }
}
