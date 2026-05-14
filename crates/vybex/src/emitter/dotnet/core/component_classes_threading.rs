use super::super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        DotnetClassExport::new(
            "dotnet.System.Threading.Tasks",
            ClassType::new("Task")
                .with_constructor(ConstructorDef::new(1).with_common_backing("threading.thread_new"))
                .with_method(MethodDef::new("Start", 0, MethodBody::Common("threading.thread_start".into())))
                .with_method(MethodDef::static_method("Run", 1, MethodBody::Common("threading.task_run".to_string())))
                .with_method(MethodDef::static_method("Delay", 1, MethodBody::Common("threading.task_delay".into()))),
        ),
        DotnetClassExport::new(
            "dotnet.System.Threading",
            ClassType::new("Interlocked")
                .with_method(MethodDef::static_method("Add", 2, MethodBody::Common("threading.atomic_add".to_string())))
                .with_method(MethodDef::static_method("Exchange", 2, MethodBody::Common("threading.atomic_xchg".to_string())))
                .with_method(MethodDef::static_method("CompareExchange", 3, MethodBody::Common("threading.atomic_cmpxchg".to_string()))),
        ),
        DotnetClassExport::new(
            "dotnet.System.Threading",
            ClassType::new("Thread")
                .with_constructor(ConstructorDef::new(1).with_common_backing("threading.thread_new"))
                .with_method(MethodDef::new("Start", 0, MethodBody::Common("threading.thread_start".into())))
                .with_method(MethodDef::new("Join", 0, MethodBody::Common("threading.thread_join".into())))
                .with_method(MethodDef::static_method("Sleep", 1, MethodBody::Common("threading.sleep".to_string()))),
        ),
    ]
}
