use std::sync::LazyLock;

use super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{
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
                    MethodBody::Common("dotnet.winforms_application_run".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Exit",
                    0,
                    MethodBody::Common("dotnet.winforms_application_exit".into()),
                ))
                .with_method(MethodDef::static_method(
                    "EnableVisualStyles",
                    0,
                    MethodBody::Common("dotnet.winforms_noop".into()),
                ))
                .with_method(MethodDef::static_method(
                    "SetCompatibleTextRenderingDefault",
                    1,
                    MethodBody::Common("dotnet.winforms_noop".into()),
                )),
        ));
        exports.push(DotnetClassExport::new(
            "dotnet.System.Windows.Forms",
            ClassType::new("MessageBox")
                .with_method(MethodDef::static_method(
                    "Show",
                    1,
                    MethodBody::Common("dotnet.winforms_message_box_show".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Show",
                    2,
                    MethodBody::Common("dotnet.winforms_message_box_show".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Show",
                    3,
                    MethodBody::Common("dotnet.winforms_message_box_show".into()),
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
        _ if matches!(
            class.name,
            "Graphics"
                | "Pen"
                | "Brush"
                | "SolidBrush"
                | "HatchBrush"
                | "LinearGradientBrush"
                | "Point"
                | "Size"
                | "SizeF"
                | "Font"
                | "Color"
        ) =>
        {
            "dotnet.System.Drawing"
        }
        _ => "dotnet.System.Windows.Forms",
    }
}

fn class_to_component_class(class: &DotnetClass) -> ClassType {
    let mut out = match class.parent {
        Some(parent) => ClassType::new(class.name).with_parent(parent),
        None => ClassType::new(class.name),
    };

    for prop in class.properties {
        out = out.with_property(
            PropertyDef::new(*prop)
                .with_setter(HostTarget::new(
                    "vybe:gui",
                    vybe_compiler::primitives::gui::HOST_FN_SET_PROPERTY,
                ))
                .with_getter(HostTarget::new(
                    "vybe:gui",
                    vybe_compiler::primitives::gui::HOST_FN_GET_PROPERTY,
                )),
        );
    }

    for method in class.methods {
        let param_count = method.arity.saturating_sub(1);
        match method.target {
            MethodTarget::Host { module, fn_name } => {
                out = out.with_method(MethodDef::new(
                    method.name,
                    param_count,
                    MethodBody::HostCall(HostTarget::new(module, fn_name)),
                ));
            }
            // `CreateGraphics` is a non-trivial method Body (construct a
            // Graphics stamped with the control's name). Model it as a shared
            // dotnet emitter so it resolves through the descriptor at the call
            // site — control leaves no longer emit a ctor chunk to bind it.
            // No-op Body methods (SuspendLayout/…) are left to the profile's
            // `noop` value-method, and other Body methods stay ctor-bound.
            MethodTarget::Body(_) if method.name == "CreateGraphics" => {
                out = out.with_method(MethodDef::new(
                    method.name,
                    param_count,
                    MethodBody::Common("dotnet.control_create_graphics".to_string()),
                ));
            }
            // Drawing Body methods (Graphics/Pen/Brush: DrawLine, FillRectangle,
            // transforms, …) resolve through the descriptor and lower inline at
            // the call site via `builder::emit_body_inline` — the drawing object
            // needs no ctor-bound thunk.
            MethodTarget::Body(_) if !method.target.is_noop() => {
                out = out.with_method(MethodDef::new(
                    method.name,
                    param_count,
                    MethodBody::Common(format!("dotnet.drawing.{}", method.name)),
                ));
            }
            // A no-op is still a DECLARED method. This used to say the profile's
            // `noop` value-method would answer instead, and nothing consumed
            // `noop_methods` any more — so `Me.SuspendLayout()` resolved to a
            // `struct.get suspendlayout` on the element, which is `undefined`,
            // and every designer-generated `InitializeComponent` died on its
            // first line. Declaring it is what makes it resolvable at all;
            // `dotnet.winforms_noop` is the emit that does nothing, and the
            // descriptor carries the arity so an override still wins.
            MethodTarget::Body(_) => {
                out = out.with_method(MethodDef::new(
                    method.name,
                    param_count,
                    MethodBody::Common("dotnet.winforms_noop".to_string()),
                ));
            }
            _ => {}
        }
    }

    if let Some(host_fn) = class.widget_host_fn {
        out = out.with_constructor(
            ConstructorDef::new(class.ctor_arity)
                .with_backing(HostTarget::new(class.widget_host_module, host_fn)),
        );
    } else if crate::emitter::tree_register::is_element_mapped(class.name) {
        // An element-mapped class is CONSTRUCTIBLE without a `vybe:gui`
        // factory: the element mapping is what materializes it, and the
        // registrar turns that into a `CtorSpec` whose `control_fn` creates
        // the node. The backing stays `None` on purpose — there is no host
        // function to call, and inventing one would put the object back on
        // the path this platform is being converted off.
        //
        // Without this, a class that has no `vybe:gui` twin gets NO
        // constructor at all and `New ToolStripMenuItem()` reaches
        // "undefined is not callable". `vybe:gui` only registers `new_*` for
        // the names in its own `control_types` list, and menu ITEMS are not
        // in it — which is precisely why they need this door rather than a
        // new entry in a host list.
        out = out.with_constructor(ConstructorDef::new(class.ctor_arity));
    }

    out
}
