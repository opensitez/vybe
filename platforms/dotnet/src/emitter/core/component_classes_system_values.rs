use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{
    ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef,
};

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
        // `char.ToUpper('q')` reached undefined and trapped — while the SAME
        // conversions already worked in Java, whose `java.lang.Character`
        // registers these leaves against `ecma:char` and `primitives::strings`.
        // Nothing was missing but the registration.
        if matches!(name, "Char" | "char") {
            for (method, host_fn) in [
                ("IsDigit", "isDigit"),
                ("IsLetter", "isLetter"),
                ("IsLetterOrDigit", "isAlnum"),
                ("IsUpper", "isUpper"),
                ("IsLower", "isLower"),
                ("IsWhiteSpace", "isSpace"),
            ] {
                ty = ty.with_method(MethodDef::static_method(
                    method,
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:char", host_fn)),
                ));
            }
            // The same host fns `primitives::strings::emit_to_upper` /
            // `emit_to_lower` call. Declared as `HostCall` rather than
            // `Common`, because `tree_register` turns a `Common` body into a
            // `CommonEmit` leaf whose name has to be one the dispatch chain
            // resolves — `strings.to_upper` is not, and the call reached
            // undefined.
            ty = ty
                .with_method(MethodDef::static_method(
                    "ToUpper",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:string", "toUpperCase")),
                ))
                .with_method(MethodDef::static_method(
                    "ToLower",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:string", "toLowerCase")),
                ));
        }
        if matches!(name, "Int32" | "int") {
            ty = ty.with_method(MethodDef::static_method(
                "TryParse",
                1,
                MethodBody::Common("dotnet.try_parse_int".into()),
            ));
        }
        exports.push(DotnetClassExport::new("dotnet.System", ty));
    }

    exports
}
