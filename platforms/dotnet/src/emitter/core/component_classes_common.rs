use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef};

pub(super) fn constructor_class(
    interface: &'static str,
    name: &'static str,
    module: &'static str,
    ctor: &'static str,
) -> DotnetClassExport {
    DotnetClassExport::new(
        interface,
        ClassType::new(name)
            .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new(module, ctor))),
    )
}

pub(super) fn common_constructor_class(
    interface: &'static str,
    name: &'static str,
    emit: &'static str,
) -> DotnetClassExport {
    DotnetClassExport::new(
        interface,
        ClassType::new(name).with_constructor(ConstructorDef::new(0).with_common_backing(emit)),
    )
}

pub(super) fn constructor_and_static_class(
    interface: &'static str,
    name: &'static str,
    constructor: Option<(&'static str, &'static str)>,
    methods: &[(&'static str, u8, &'static str, &'static str)],
) -> DotnetClassExport {
    let mut class = ClassType::new(name);
    if let Some((module, ctor)) = constructor {
        class = class
            .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new(module, ctor)));
    }
    for (method, arity, module, func) in methods {
        class = class.with_method(MethodDef::static_method(
            *method,
            *arity,
            MethodBody::HostCall(HostTarget::new(*module, *func)),
        ));
    }
    DotnetClassExport::new(interface, class)
}

pub(super) fn static_only_class(
    interface: &'static str,
    name: &'static str,
    methods: &[(&'static str, u8, &'static str, &'static str)],
) -> DotnetClassExport {
    constructor_and_static_class(interface, name, None, methods)
}
