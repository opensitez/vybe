use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![DotnetClassExport::new(
        "dotnet.System",
        ClassType::new("Version")
            .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.version_new"))
            .with_constructor(ConstructorDef::new(2).with_common_backing("dotnet.version_new"))
            .with_constructor(ConstructorDef::new(3).with_common_backing("dotnet.version_new"))
            .with_constructor(ConstructorDef::new(4).with_common_backing("dotnet.version_new"))
            .with_method(MethodDef::static_method(
                "Parse",
                1,
                MethodBody::Common("dotnet.version_parse".into()),
            ))
            .with_method(MethodDef::static_method(
                "TryParse",
                2,
                MethodBody::Common("dotnet.version_try_parse".into()),
            ))
            .with_method(MethodDef::static_method(
                "CompareTo",
                2,
                MethodBody::Common("dotnet.version_compare_instance".into()),
            ))
            .with_method(MethodDef::static_method(
                "Equals",
                2,
                MethodBody::Common("dotnet.version_equals".into()),
            ))
            // The relational operators. `version_adapter` has had
            // `emit_version_lt`/`emit_version_gt` and their `dotnet.version_lt`/
            // `dotnet.version_gt` dispatch keys all along, but NO LEAF declared
            // them — so nothing could resolve to them and a consumer had to
            // rebuild the ordering itself. Registered under .NET's own operator
            // method names so the spelling is the framework's, not ours.
            .with_method(MethodDef::static_method(
                "op_LessThan",
                2,
                MethodBody::Common("dotnet.version_lt".into()),
            ))
            .with_method(MethodDef::static_method(
                "op_GreaterThan",
                2,
                MethodBody::Common("dotnet.version_gt".into()),
            ))
            .with_method(MethodDef::static_method(
                "op_Equality",
                2,
                MethodBody::Common("dotnet.version_equals".into()),
            ))
            .with_method(MethodDef::new(
                "ToString",
                0,
                MethodBody::Common("dotnet.version_to_string".into()),
            ))
            .with_method(MethodDef::new(
                "ToString",
                1,
                MethodBody::Common("dotnet.version_to_string".into()),
            ))
            .with_method(MethodDef::new(
                "Clone",
                0,
                MethodBody::Common("dotnet.version_clone".into()),
            ))
            .with_method(MethodDef::new(
                "Equals",
                1,
                MethodBody::Common("dotnet.version_equals".into()),
            ))
            .with_method(MethodDef::new(
                "CompareTo",
                1,
                MethodBody::Common("dotnet.version_compare_instance".into()),
            )),
    )]
}
