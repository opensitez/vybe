//! `Form` — the WinForms top-level window.
//!
//! `Form : ContainerControl : ScrollableControl : Control : Component :
//! MarshalByRefObject : Object`. The first concrete leaf in the .NET
//! hierarchy. Its constructor calls NO host factory: `Form` is element-mapped
//! to `body`, and the element mapping is what materialises it
//! (`tree_register::is_element_mapped`), after the parent chain has bound the
//! inherited control / scrollable / container setters.
//!
//! A form is a WINDOW, so its verbs are window verbs: `Close` and `Dispose`
//! are `window.close()`, `Activate` is `window.focus()`, and `CenterToScreen`
//! is `moveTo` — each reached through `activeDocument().defaultView`, because a
//! guest has no global object to spell `window` with.
//!
//! ## Property placement
//!
//! `Text` is **inherited from Control**, not redeclared here. In real
//! WinForms `Form.Text` is the window title and `Control.Text` is the
//! generic display text — they share the same `Text` property because
//! Form just overrides `WindowText` semantics at the OS level. We get the
//! same effect for free: the `__set_text` setter installed by `Control`
//! mirrors to `gui.set_property(form_name, "Text", v)`, and the gui state
//! registry treats that as the form title because the form widget reads
//! its title from the `Text` key.
//!
//! Properties listed here are **only** the ones .NET adds at the `Form`
//! class level (per MSDN's `Form` class members page).

use super::{DotnetClass, DotnetMethod, MethodTarget};

/// Methods owned by `Form`. `Show`/`Hide`/`Focus` are inherited from
/// `Control` and don't need to be re-listed here. `Close`/`ShowDialog`/
/// `Activate`/`CenterToScreen` are form-specific.
const FORM_METHODS: &[DotnetMethod] = &[
    DotnetMethod {
        name: "Close",
        arity: 1,
        target: MethodTarget::common("dotnet.winforms_form_close"),
    },
    // A form's disposal is its WINDOW closing, not its node being removed.
    //
    // ⛔ This declaration is what makes `Control.Dispose` safe to point at the
    // DOM verb. `Form` maps to `body`, so `element.remove()` on a form would
    // delete the document body and take the whole app shell with it — which is
    // exactly why `gui_control_verb` guards `dispose` with `!is_form`. That
    // guard only sends Form to the INHERITED target; it cannot say what the
    // target should be. Declaring `Dispose` here does, and `inherited_methods`
    // yields nearest-first so this shadows `Control`'s.
    DotnetMethod {
        name: "Dispose",
        arity: 1,
        target: MethodTarget::common("dotnet.winforms_form_close"),
    },
    DotnetMethod {
        name: "ShowDialog",
        arity: 1,
        target: MethodTarget::common("dotnet.winforms_form_show_dialog"),
    },
    DotnetMethod {
        name: "Activate",
        arity: 1,
        target: MethodTarget::common("dotnet.winforms_form_activate"),
    },
    DotnetMethod {
        name: "CenterToScreen",
        arity: 1,
        target: MethodTarget::common("dotnet.winforms_form_center_to_screen"),
    },
];

pub fn classes() -> &'static [DotnetClass] {
    &[DotnetClass {
        name: "Form",
        parent: Some("ContainerControl"),
        properties: &[
            // Window chrome
            "FormBorderStyle",
            "ControlBox",
            "MaximizeBox",
            "MinimizeBox",
            "HelpButton",
            "ShowIcon",
            "Icon",
            // Position / state
            "StartPosition",
            "WindowState",
            "TopMost",
            "Opacity",
            // Sizing
            "AutoSize",
            "AutoSizeMode",
            "AutoScaleMode",
            "AutoScaleDimensions",
            // Behaviour
            "ShowInTaskbar",
            "KeyPreview",
            "AcceptButton",
            "CancelButton",
            "DialogResult",
            // Owner / parent
            "Owner",
            "MdiParent",
            "IsMdiContainer",
            // Menu / status
            "MainMenuStrip",
            // Misc
            "TransparencyKey",
        ],
        methods: FORM_METHODS,
        ctor_arity: 0,
        // A Form IS the document. `html_element_for_control` maps it to `body`
        // and `emit_control_element` special-cases it to `document.body`, so
        // construction has been element-shaped all along for a SUBCLASS —
        // `class Form1 : Form` went through the control path. Only a bare
        // `new Form()` still reached the factory.
        //
        // What kept the factory alive was the `Controls` collection it
        // installed as a field. `Controls` is now declared on every
        // element-backed control and answers `dotnet.self` — a form's children
        // ARE its element's children — so the factory has nothing left to
        // provide.
        widget_host_fn: None,    }]
}
