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

/// A descriptor class whose CONSTRUCTOR is composed from primitives rather than
/// called on a host.
///
/// The first arm of `class_to_component_class`'s constructor gate: the element
/// mapping materialises a node, and this builds the value in bytecode with no
/// host at all.
///
/// A value type is four numbers in an object — nothing a host is needed for —
/// so `Rectangle` is built by `dotnet.rectangle_new` (`dispatch.rs`) the way
/// `StringBuilder` is built by `dotnet.string_builder_new`.
pub(crate) fn common_ctor_for(class: &str) -> Option<&'static str> {
    match class {
        "Rectangle" => Some("dotnet.rectangle_new"),
        // A value type is not a widget — there is no element and nothing to
        // insert — so the object is composed here.
        "Point" => Some("dotnet.point_new"),
        "Size" => Some("dotnet.size_new"),
        "Font" => Some("dotnet.font_new"),
        "Color" => Some("dotnet.color_new"),
        "Pen" => Some("dotnet.pen_new"),
        "SolidBrush" => Some("dotnet.solid_brush_new"),
        "Graphics" => Some("dotnet.graphics_new"),
        "HatchBrush" => Some("dotnet.hatch_brush_new"),
        "LinearGradientBrush" => Some("dotnet.linear_gradient_brush_new"),
        _ => None,
    }
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

/// Not a host module — a MARKER.
///
/// A property carries a `HostTarget` because that is the shape `PropertyDef`
/// takes, and `tree_register::accessor_node` reads the target's NAME to
/// recognise a generic control accessor and rewrite it into the shared role
/// emits (`gui.prop_set.<role>` → `web:dom`/`web:cssom`). The module half is
/// never imported and never called.
///
/// ⚠ It must never name a real module. A marker that looks like a host is read
/// as one by the next person, and by grep.
const PROPERTY_MARKER: &str = "dotnet.property-marker";

fn class_to_component_class(class: &DotnetClass) -> ClassType {
    let mut out = match class.parent {
        Some(parent) => ClassType::new(class.name).with_parent(parent),
        None => ClassType::new(class.name),
    };

    for prop in class.properties {
        out = out.with_property(
            PropertyDef::new(*prop)
                .with_setter(HostTarget::new(
                    PROPERTY_MARKER,
                    vybe_compiler::primitives::gui::HOST_FN_SET_PROPERTY,
                ))
                .with_getter(HostTarget::new(
                    PROPERTY_MARKER,
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
            // A shared emit, declared by name. This arm is written out rather
            // than left to the `_ => {}` below on purpose: a variant that
            // falls through there declares NO method at all, and an undeclared
            // method is not a compile error — it is `Me.Focus()` resolving to
            // `undefined`, which is the failure the comment further down
            // records for `SuspendLayout`.
            MethodTarget::Common { emit } => {
                out = out.with_method(MethodDef::new(
                    method.name,
                    param_count,
                    MethodBody::Common(emit.to_string()),
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

    if let Some(emit) = common_ctor_for(class.name) {
        out = out
            .with_constructor(ConstructorDef::new(class.ctor_arity).with_common_backing(emit));
    } else if crate::emitter::tree_register::is_element_mapped(class.name) {
        // An element-mapped class is CONSTRUCTIBLE without any host factory:
        // the element mapping is what materializes it, and the
        // registrar turns that into a `CtorSpec` whose `control_fn` creates
        // the node. The backing stays `None` on purpose — there is no host
        // function to call, and inventing one would put the object back on a
        // path this platform does not have any more.
        //
        // Without this, a class gets NO constructor at all and
        // `New ToolStripMenuItem()` reaches "undefined is not callable" — menu
        // ITEMS are elements like any other control, and this is the door that
        // says so.
        out = out.with_constructor(ConstructorDef::new(class.ctor_arity));
    }

    out
}
