use super::super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    let mut exports = vec![DotnetClassExport::new(
        "dotnet.System",
        ClassType::new("Guid")
            .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.guid_new"))
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
