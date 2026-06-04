use super::super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![DotnetClassExport::new(
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
    )]
}
