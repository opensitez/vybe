use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    let mut exports = vec![DotnetClassExport::new(
        "dotnet.System",
        ClassType::new("Guid")
            .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.guid_new"))
            .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.guid_new"))
            .with_method(MethodDef::static_method(
                "Empty",
                0,
                MethodBody::Common("dotnet.guid_empty".into()),
            ))
            .with_method(MethodDef::static_method(
                "NewGuid",
                0,
                MethodBody::Common("dotnet.guid_new_guid".into()),
            ))
            .with_method(MethodDef::static_method(
                "Parse",
                1,
                MethodBody::Common("dotnet.guid_parse".into()),
            ))
            .with_method(MethodDef::static_method(
                "TryParse",
                2,
                MethodBody::Common("dotnet.guid_try_parse".into()),
            ))
            .with_method(MethodDef::new(
                "ToString",
                0,
                MethodBody::Common("dotnet.guid_to_string".into()),
            ))
            .with_method(MethodDef::new(
                "ToString",
                1,
                MethodBody::Common("dotnet.guid_to_string".into()),
            ))
            .with_method(MethodDef::new(
                "ToByteArray",
                0,
                MethodBody::Common("dotnet.guid_to_byte_array".into()),
            ))
            .with_method(MethodDef::new(
                "GetHashCode",
                0,
                MethodBody::Common("dotnet.guid_get_hash_code".into()),
            )),
    )];

    for (name, parse_emit) in [
        ("Int32", "dotnet.parse_int"),
        ("int", "dotnet.parse_int"),
        ("Byte", "dotnet.parse_byte"),
        ("byte", "dotnet.parse_byte"),
        ("Int64", "dotnet.parse_long"),
        ("long", "dotnet.parse_long"),
        ("Single", "dotnet.parse_float"),
        ("float", "dotnet.parse_float"),
        ("Decimal", "dotnet.parse_decimal"),
        ("decimal", "dotnet.parse_decimal"),
        ("Double", "dotnet.parse_double"),
        ("double", "dotnet.parse_double"),
        ("Boolean", "dotnet.parse_bool"),
        ("bool", "dotnet.parse_bool"),
        ("Char", "dotnet.parse_char"),
        ("char", "dotnet.parse_char"),
    ] {
        let mut ty = ClassType::new(name).with_method(MethodDef::static_method(
            "Parse",
            1,
            MethodBody::Common(parse_emit.into()),
        ));
        if matches!(name, "Decimal" | "decimal") {
            ty = ty
                .with_method(MethodDef::static_method(
                    "Round",
                    1,
                    MethodBody::Common("dotnet.system.math.round".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Round",
                    2,
                    MethodBody::Common("dotnet.system.math.round".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Round",
                    3,
                    MethodBody::Common("dotnet.system.math.round".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Truncate",
                    1,
                    MethodBody::Common("dotnet.system.math.truncate".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Floor",
                    1,
                    MethodBody::Common("dotnet.system.math.floor".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Ceiling",
                    1,
                    MethodBody::Common("dotnet.system.math.ceiling".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Abs",
                    1,
                    MethodBody::Common("dotnet.system.math.abs".into()),
                ));
        }
        // `System.Char`'s static surface. Only `Parse` was registered, so
        // `char.ToUpper('q')` reached undefined and trapped.
        //
        // There is no `ecma:char` host module and there never was — ECMAScript
        // has no char type, so a char IS a one-character string, and jvm/java
        // both carry a note that binding to `ecma:char` panicked at compile
        // time. The classifiers are the SHARED string primitives
        // (`primitives::strings`, pure WASM over code points); case conversion
        // is `ecma:string`, which is a real ECMA String method.
        if matches!(name, "Char" | "char") {
            // ⛔ Static lookup (`emitter/mod.rs:574`) matches on NAME ALONE and
            // never reads `arity` — unlike the instance path right above it.
            // So the six classifiers below cannot be given a second arity-2
            // registration for the `(string, index)` overload: the first row
            // would win and the second would be dead code. Each now routes to
            // `char_adapter`, whose body branches on `argc`.
            //
            // Before this, `Char.IsDigit(text, 1)` reached the arity-1
            // `str_is_digit` with an extra value on the stack and trapped in
            // `wasm:js-string.length — not a string`.
            for (method, emit) in [
                ("IsDigit", "dotnet.char_is_digit"),
                ("IsLetter", "dotnet.char_is_letter"),
                ("IsLetterOrDigit", "dotnet.char_is_letter_or_digit"),
                ("IsUpper", "dotnet.char_is_upper"),
                ("IsLower", "dotnet.char_is_lower"),
                ("IsWhiteSpace", "dotnet.char_is_white_space"),
                ("ToUpper", "dotnet.char_to_upper"),
                ("ToLower", "dotnet.char_to_lower"),
                // The rest of the static surface, registered nowhere before —
                // every one of these resolved to nothing and rendered empty.
                ("IsAscii", "dotnet.char_is_ascii"),
                ("IsAsciiDigit", "dotnet.char_is_ascii_digit"),
                ("IsAsciiLetter", "dotnet.char_is_ascii_letter"),
                ("IsAsciiLetterOrDigit", "dotnet.char_is_ascii_letter_or_digit"),
                ("IsAsciiHexDigit", "dotnet.char_is_ascii_hex_digit"),
                ("IsControl", "dotnet.char_is_control"),
                ("IsSeparator", "dotnet.char_is_separator"),
                ("IsPunctuation", "dotnet.char_is_punctuation"),
                ("IsSymbol", "dotnet.char_is_symbol"),
                ("IsSurrogate", "dotnet.char_is_surrogate"),
                ("IsHighSurrogate", "dotnet.char_is_high_surrogate"),
                ("IsLowSurrogate", "dotnet.char_is_low_surrogate"),
                ("IsSurrogatePair", "dotnet.char_is_surrogate_pair"),
                ("ConvertToUtf32", "dotnet.char_convert_to_utf32"),
                ("ConvertFromUtf32", "dotnet.char_convert_from_utf32"),
                ("GetNumericValue", "dotnet.char_get_numeric_value"),
                ("GetUnicodeCategory", "dotnet.char_get_unicode_category"),
            ] {
                ty = ty.with_method(MethodDef::static_method(
                    method,
                    1,
                    MethodBody::Common(emit.into()),
                ));
            }
        }
        // ⛔ The 1-arg `TryParse` EVERY numeric type needs.
        // `lowering::try_parse_desugar` rewrites `T.TryParse(s, out)` into
        // `(out = T.TryParse(s)) <> Nothing`, so the one-argument form is the
        // core of the whole feature — and only `Int32` had it. `Double`,
        // `Single`, `Int64`, `Byte` and `Decimal` resolved it to nothing and
        // answered `null`.
        //
        // Integral types floor, so they take the `int` body; the rest keep
        // their fraction.
        let try_parse_emit = match name {
            "Int32" | "int" | "Int64" | "long" | "Byte" | "byte" | "Int16" | "short" => {
                Some("dotnet.try_parse_int")
            }
            "Single" | "float" | "Double" | "double" | "Decimal" | "decimal" => {
                Some("dotnet.try_parse_double")
            }
            _ => None,
        };
        if let Some(emit) = try_parse_emit {
            ty = ty.with_method(MethodDef::static_method(
                "TryParse",
                1,
                MethodBody::Common(emit.into()),
            ));
        }
        // ⛔ The float predicates belong on THIS registration, not a second
        // `ClassType::new("Double")` of their own. Two exports of one class
        // name under one interface do not merge — the class lookup takes the
        // first match — so declaring `IsNaN` separately left it unreachable
        // while `Parse` on the other copy resolved. It also gets them the
        // lowercase spellings (`double.IsNaN`) for free.
        if matches!(name, "Single" | "float" | "Double" | "double") {
            for (member, emit) in [
                ("IsNaN", "dotnet.double_is_nan"),
                ("IsInfinity", "dotnet.double_is_infinity"),
                ("IsPositiveInfinity", "dotnet.double_is_positive_infinity"),
                ("IsNegativeInfinity", "dotnet.double_is_negative_infinity"),
                ("IsFinite", "dotnet.double_is_finite"),
                ("IsNormal", "dotnet.double_is_normal"),
                ("IsSubnormal", "dotnet.double_is_subnormal"),
            ] {
                ty = ty.with_method(MethodDef::static_method(
                    member,
                    1,
                    MethodBody::Common(emit.into()),
                ));
            }
        }
        exports.push(DotnetClassExport::new("dotnet.System", ty));
    }

    exports
}
