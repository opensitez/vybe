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

use super::DotnetClass;

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
            widget_host_fn: None,
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
            widget_host_fn: None,
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
            widget_host_fn: None,
        },
    ]
}
