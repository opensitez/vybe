use super::super::super::class_exports::DotnetClassExport;
use super::component_classes_common::{common_constructor_class, constructor_and_static_class, static_only_class};
use vybe_bytecode::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        constructor_and_static_class(
            "dotnet.System.Diagnostics",
            "Stopwatch",
            Some(("wasi:clocks", "stopwatchNew")),
            &[("StartNew", 0, "wasi:clocks", "stopwatchNew")],
        ),
        static_only_class(
            "dotnet.System.Diagnostics",
            "Debug",
            &[
                ("WriteLine", 1, "wasi:cli", "log"),
                ("Write", 1, "wasi:cli", "log"),
                ("Assert", 1, "wasi:cli", "log"),
            ],
        ),
        static_only_class(
            "dotnet.System.Diagnostics",
            "Trace",
            &[("WriteLine", 1, "wasi:cli", "log")],
        ),
        common_constructor_class(
            "dotnet.System.Diagnostics",
            "ProcessStartInfo",
            "dotnet.process_start_info_new",
        ),
        DotnetClassExport::new(
            "dotnet.System.Diagnostics",
            ClassType::new("Process")
                .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.process_new"))
                .with_method(MethodDef::static_method("Start", 1, MethodBody::Common("dotnet.process_start".into())))
                .with_method(MethodDef::static_method("GetCurrentProcess", 0, MethodBody::Common("dotnet.process_get_current".into())))
                .with_method(MethodDef::new("WaitForExit", 0, MethodBody::Common("dotnet.process_wait_for_exit".into()))),
        ),
    ]
}
