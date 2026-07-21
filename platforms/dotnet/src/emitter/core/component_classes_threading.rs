use super::super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    let mut task = ClassType::new("Task")
        .with_constructor(ConstructorDef::new(1).with_common_backing("threading.thread_new"))
        .with_method(MethodDef::new(
            "Start",
            0,
            MethodBody::Common("threading.thread_start".into()),
        ))
        .with_method(MethodDef::new(
            "Wait",
            0,
            MethodBody::Common("dotnet.task_wait".into()),
        ))
        .with_method(MethodDef::static_method(
            "Run",
            1,
            MethodBody::Common("threading.task_run".to_string()),
        ))
        .with_method(MethodDef::static_method(
            "Delay",
            1,
            MethodBody::Common("threading.task_delay".into()),
        ))
        .with_method(MethodDef::static_method(
            "FromResult",
            1,
            MethodBody::Common("dotnet.task_from_result".into()),
        ))
        .with_method(MethodDef::new(
            "ContinueWith",
            1,
            MethodBody::Common("dotnet.task_continue_with".into()),
        ));
    // `Task.WhenAll` / `WhenAny` take variadic tasks; overload resolution is by
    // exact arity, so register each width plus the single `IEnumerable` form.
    for n in 1..=8u8 {
        task = task
            .with_method(MethodDef::static_method(
                "WhenAll",
                n,
                MethodBody::Common("dotnet.task_when_all".into()),
            ))
            .with_method(MethodDef::static_method(
                "WhenAny",
                n,
                MethodBody::Common("dotnet.task_when_any".into()),
            ));
    }
    vec![
        DotnetClassExport::new("dotnet.System.Threading.Tasks", task),
        DotnetClassExport::new(
            "dotnet.System.Threading",
            ClassType::new("Interlocked")
                .with_method(MethodDef::static_method(
                    "Add",
                    2,
                    MethodBody::Common("threading.atomic_add".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "Exchange",
                    2,
                    MethodBody::Common("threading.atomic_xchg".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "CompareExchange",
                    3,
                    MethodBody::Common("threading.atomic_cmpxchg".to_string()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.Threading",
            ClassType::new("Thread")
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("threading.thread_new"),
                )
                .with_method(MethodDef::new(
                    "Start",
                    0,
                    MethodBody::Common("threading.thread_start".into()),
                ))
                .with_method(MethodDef::new(
                    "Join",
                    0,
                    MethodBody::Common("threading.thread_join".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Sleep",
                    1,
                    MethodBody::Common("threading.sleep".to_string()),
                )),
        ),
    ]
}
