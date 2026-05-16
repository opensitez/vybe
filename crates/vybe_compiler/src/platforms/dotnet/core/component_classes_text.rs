use super::super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        DotnetClassExport::new(
            "dotnet.System.Text",
            ClassType::new("StringBuilder")
                .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.string_builder_new"))
                .with_method(MethodDef::new("Append", 1, MethodBody::Common("dotnet.sb_append".into())))
                .with_method(MethodDef::new("AppendLine", 1, MethodBody::Common("dotnet.sb_append_line".into())))
                .with_method(MethodDef::new("ToString", 0, MethodBody::Common("dotnet.sb_to_string".into())))
                .with_method(MethodDef::new("Clear", 0, MethodBody::Common("dotnet.sb_clear".into())))
                .with_method(MethodDef::new("Length", 0, MethodBody::Common("dotnet.sb_length".into())))
                .with_method(MethodDef::new("Insert", 2, MethodBody::Common("dotnet.sb_insert".into())))
                .with_method(MethodDef::new("Remove", 2, MethodBody::Common("dotnet.sb_remove".into())))
                .with_method(MethodDef::new("Replace", 2, MethodBody::Common("dotnet.sb_replace".into()))),
        ),
        DotnetClassExport::new(
            "dotnet.System.Text.RegularExpressions",
            ClassType::new("Regex")
                .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.regex_new"))
                .with_method(MethodDef::new("IsMatch", 1, MethodBody::Common("dotnet.regex_is_match".into())))
                .with_method(MethodDef::new("Replace", 2, MethodBody::Common("dotnet.regex_replace".into())))
                .with_method(MethodDef::new("Split", 1, MethodBody::Common("dotnet.regex_split".into())))
                .with_method(MethodDef::new("Match", 1, MethodBody::Common("dotnet.regex_match".into())))
                .with_method(MethodDef::new("Matches", 1, MethodBody::Common("dotnet.regex_matches".into()))),
        ),
    ]
}
