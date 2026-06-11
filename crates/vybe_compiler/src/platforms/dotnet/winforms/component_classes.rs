use std::sync::LazyLock;

use super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{
    ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef, PropertyDef,
};

use super::classes::{self, DotnetClass, MethodTarget};

pub fn class_exports() -> &'static [DotnetClassExport] {
    static EXPORTS: LazyLock<Vec<DotnetClassExport>> = LazyLock::new(|| {
        let mut exports = Vec::new();
        for class in classes::dotnet_classes() {
            exports.push(class_to_export(class));
        }
        exports.push(DotnetClassExport::new(
            "dotnet.System.Windows.Forms",
            ClassType::new("Application")
                .with_method(MethodDef::static_method(
                    "Run",
                    1,
                    MethodBody::HostCall(HostTarget::new(
                        "vybe:gui",
                        crate::emitter::gui::HOST_FN_RUN_APPLICATION,
                    )),
                ))
                .with_method(MethodDef::static_method(
                    "Exit",
                    0,
                    MethodBody::HostCall(HostTarget::new(
                        "vybe:gui",
                        crate::emitter::gui::HOST_FN_APP_EXIT,
                    )),
                )),
        ));
        exports
    });
    EXPORTS.as_slice()
}

pub fn component_class_exports() -> &'static [(&'static str, ClassType)] {
    static EXPORTS: LazyLock<Vec<(&'static str, ClassType)>> = LazyLock::new(|| {
        class_exports()
            .iter()
            .map(|export| (export.interface, export.class.clone()))
            .collect()
    });
    EXPORTS.as_slice()
}

fn class_to_export(class: &DotnetClass) -> DotnetClassExport {
    DotnetClassExport::with_wrapper(
        dotnet_interface_for_class(class),
        class_to_component_class(class),
        *class,
    )
}

fn dotnet_interface_for_class(class: &DotnetClass) -> &'static str {
    match class.name {
        "Object" | "MarshalByRefObject" | "Component" => "dotnet.System",
        _ if matches!(class.name, "Graphics" | "Pen" | "Brush" | "SolidBrush" | "HatchBrush"
            | "LinearGradientBrush" | "Point" | "Size" | "SizeF" | "Font" | "Color") =>
            "dotnet.System.Drawing",
        _ => "dotnet.System.Windows.Forms",
    }
}

fn class_to_component_class(class: &DotnetClass) -> ClassType {
    let mut out = match class.parent {
        Some(parent) => ClassType::new(class.name).with_parent(parent),
        None => ClassType::new(class.name),
    };

    for prop in class.properties {
        out = out.with_property(PropertyDef::new(*prop).with_setter(HostTarget::new(
            "vybe:gui",
            crate::emitter::gui::HOST_FN_SET_PROPERTY,
        )));
    }

    for method in class.methods {
        if let MethodTarget::Host { module, fn_name } = method.target {
            let param_count = if method.arity > 0 {
                method.arity - 1
            } else {
                0
            };
            out = out.with_method(MethodDef::new(
                method.name,
                param_count,
                MethodBody::HostCall(HostTarget::new(module, fn_name)),
            ));
        }
    }

    if let Some(host_fn) = class.widget_host_fn {
        out = out.with_constructor(
            ConstructorDef::new(class.ctor_arity)
                .with_backing(HostTarget::new(class.widget_host_module, host_fn)),
        );
    }

    out
}
