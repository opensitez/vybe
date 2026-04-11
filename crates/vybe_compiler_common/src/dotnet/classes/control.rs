//! `Control → ScrollableControl → ContainerControl`.
//!
//! These are the abstract intermediate bases between `Component` and the
//! concrete WinForms controls. Real .NET puts the bulk of the
//! position/size/colour/behaviour properties on `Control`, autoscroll on
//! `ScrollableControl`, and active-control tracking on `ContainerControl`.
//!
//! All three are abstract — none has a `widget_host_fn`. Concrete leaves
//! (`Form`, `Button`, `Label`, …) inherit from one of them and supply the
//! widget at the bottom of their own ctor.
//!
//! ## Property placement
//!
//! Property names are PascalCase to match what `controlSetProperty`
//! receives as the gui-state-registry key. The setter is bound under
//! `__set_<lowercased>` because the canonical AST lowercases struct field
//! names — so `Me.Text = "X"` lowercases to `text` and the VM dispatches
//! `__set_text`. The two casings are kept consistent by
//! [`super::builder::build_setter_chunk`].

use super::{DotnetClass, DotnetMethod, MethodTarget};

/// Methods owned by `Control` (inherited by every concrete control,
/// including `Form`). Each entry maps to either a host fn or, when the
/// method returns another .NET class instance, to that class's ctor.
///
/// - `Show`/`Hide`/`Focus`/`Refresh`/… → host fns in `vybe:gui` (no-ops
///   in non-display contexts, real implementations under a future GUI
///   backend)
/// - `CreateGraphics` → calls the `Graphics` dotnet class ctor so the
///   returned instance has all `Graphics` methods (`DrawLine`, etc.)
///   bound on it. Going through the raw `vybe:drawing::graphicsNew`
///   host fn would skip method binding and break user code that does
///   `g.DrawLine(...)`.
///
/// Methods specific to `Form` (`Close`, `ShowDialog`, `Activate`,
/// `CenterToScreen`) live on the `Form` class so subclasses inherit them
/// only when inheriting from `Form`.
const CONTROL_METHODS: &[DotnetMethod] = &[
    DotnetMethod { name: "Show",           arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_show") },
    DotnetMethod { name: "Hide",           arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_hide") },
    DotnetMethod { name: "Focus",          arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_focus") },
    DotnetMethod { name: "Refresh",        arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_refresh") },
    DotnetMethod { name: "Invalidate",     arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_invalidate") },
    DotnetMethod { name: "Update",         arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_update") },
    DotnetMethod { name: "BringToFront",   arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_bring_to_front") },
    DotnetMethod { name: "SendToBack",     arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_send_to_back") },
    DotnetMethod { name: "Select",         arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_focus") },
    DotnetMethod { name: "Dispose",        arity: 1, target: MethodTarget::host("vybe:gui", "__ctrl_dispose") },
    DotnetMethod { name: "CreateGraphics", arity: 1, target: MethodTarget::dotnet_ctor("Graphics") },
];

pub fn classes() -> &'static [DotnetClass] {
    &[
        // ── Control ────────────────────────────────────────────────────────
        // The big one. Every visible WinForms control inherits from this
        // (transitively). Owns the position/size/colour/behaviour surface.
        DotnetClass {
            name: "Control",
            parent: Some("Component"),
            properties: &[
                // Identity
                "Name",
                "Tag",
                // Text content (overridden semantically by Label/TextBox/etc.
                // at runtime via the same setter — that's how WinForms works)
                "Text",
                // Position
                "Left",
                "Top",
                "Location",
                // Size
                "Width",
                "Height",
                "Size",
                "ClientSize",
                "MinimumSize",
                "MaximumSize",
                // Layout
                "Anchor",
                "Dock",
                "Margin",
                "Padding",
                // Visibility / interactivity
                "Visible",
                "Enabled",
                "TabIndex",
                "TabStop",
                // Colour & font
                "BackColor",
                "ForeColor",
                "Font",
                "BackgroundImage",
                "BackgroundImageLayout",
                // Cursor / context menu
                "Cursor",
                "ContextMenuStrip",
                // Misc
                "AllowDrop",
                "RightToLeft",
                "AccessibleName",
                "AccessibleDescription",
                "AccessibleRole",
            ],
            methods: CONTROL_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        // ── ScrollableControl ──────────────────────────────────────────────
        // Adds the autoscroll surface used by Form, Panel, …
        DotnetClass {
            name: "ScrollableControl",
            parent: Some("Control"),
            properties: &[
                "AutoScroll",
                "AutoScrollPosition",
                "AutoScrollMargin",
                "AutoScrollMinSize",
                "HScroll",
                "VScroll",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        // ── ContainerControl ───────────────────────────────────────────────
        // Adds the active-control / parent-form tracking used by Form,
        // UserControl, …
        DotnetClass {
            name: "ContainerControl",
            parent: Some("ScrollableControl"),
            properties: &[
                "ActiveControl",
                "ParentForm",
                "AutoValidate",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
    ]
}
