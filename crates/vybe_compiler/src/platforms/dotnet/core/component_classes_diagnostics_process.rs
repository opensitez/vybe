use super::super::super::class_exports::DotnetClassExport;
use super::component_classes_common::{common_constructor_class, static_only_class};
use vybe_bytecode::component_model::{
    ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef,
};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        DotnetClassExport::new(
            "dotnet.System.Diagnostics",
            ClassType::new("Stopwatch")
                .with_constructor(
                    ConstructorDef::new(0)
                        .with_backing(HostTarget::new("wasi:clocks", "stopwatchNew")),
                )
                .with_method(MethodDef::static_method(
                    "StartNew",
                    0,
                    MethodBody::Common("dotnet.stopwatch_start_new".into()),
                ))
                .with_method(MethodDef::new(
                    "Start",
                    0,
                    MethodBody::HostCall(HostTarget::new("wasi:clocks", "stopwatchStart")),
                ))
                .with_method(MethodDef::new(
                    "Stop",
                    0,
                    MethodBody::HostCall(HostTarget::new("wasi:clocks", "stopwatchStop")),
                ))
                .with_method(MethodDef::new(
                    "Reset",
                    0,
                    MethodBody::HostCall(HostTarget::new("wasi:clocks", "stopwatchReset")),
                ))
                .with_method(MethodDef::new(
                    "Restart",
                    0,
                    MethodBody::Common("dotnet.stopwatch_restart".into()),
                ))
                .with_method(MethodDef::new(
                    "ElapsedMilliseconds",
                    0,
                    MethodBody::Common("dotnet.stopwatch_elapsed_ms".into()),
                ))
                .with_method(MethodDef::new(
                    "IsRunning",
                    0,
                    MethodBody::Common("dotnet.stopwatch_is_running".into()),
                )),
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
                .with_method(MethodDef::static_method(
                    "Start",
                    1,
                    MethodBody::Common("dotnet.process_start".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetCurrentProcess",
                    0,
                    MethodBody::Common("dotnet.process_get_current".into()),
                ))
                .with_method(MethodDef::new(
                    "WaitForExit",
                    0,
                    MethodBody::Common("dotnet.process_wait_for_exit".into()),
                )),
        ),
    ]
}
