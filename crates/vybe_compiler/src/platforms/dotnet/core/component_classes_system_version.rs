use super::super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        DotnetClassExport::new(
            "dotnet.System",
            ClassType::new("Version")
                .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.version_new"))
                .with_method(MethodDef::static_method("Parse", 1, MethodBody::Common("dotnet.version_parse".into())))
                .with_method(MethodDef::new("ToString", 0, MethodBody::Common("dotnet.version_to_string".into())))
                .with_method(MethodDef::new("Equals", 1, MethodBody::Common("dotnet.version_equals".into())))
                .with_method(MethodDef::new("CompareTo", 1, MethodBody::Common("dotnet.version_compare".into()))),
        ),
    ]
}