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
//! `__set_text`. Property names here are registered onto the class's tree
//! `Type` node as `Property { get, set }` members by
//! `emitter::tree_register`, which owns the casing.

use super::{DotnetClass, DotnetMethod, MethodOp, MethodTarget};

const CONTROL_NOOP: &[MethodOp] = &[MethodOp::PushConstNull, MethodOp::Return];

/// `Control.CreateGraphics()` body.
///
/// Translation:
///
/// ```text
/// vybe:gui::createGraphics(this.__control_name) → Graphics handle
/// ```
///
/// The handle is a small Object stamped with `__type = "Graphics"` and
/// `__control_name = <this control's name>`. Subsequent canvas calls
/// (issued by Graphics method bodies) read the name out of the handle
/// to find the target `RecordingCanvas` on `GuiState`.
///
/// Note: the returned handle is NOT a fully-bound dotnet `Graphics`
/// instance — it doesn't have the `DrawLine`/`FillRectangle` method
/// thunks bound on it. That's intentional and correct: the dotnet
/// method dispatch path looks up `Graphics` methods on the inheritance
/// chain via type registry / __tid_, not via per-instance struct field
/// reads. The handle just needs to carry the canvas's identity.
///
/// (Wait, that's wrong — methods ARE bound per-instance via struct_set
/// in the ctor. We need to either go through the Graphics dotnet ctor
/// OR bind the methods on the handle here. Going through the dotnet
/// ctor IS the cleanest solution, and we already have a `NewDotnet`
/// op... but it discards args. Need to think again.)
///
/// Resolution: emit `NewDotnet { class: "Graphics", argc: 0 }` to
/// produce a fully-bound Graphics instance, THEN stamp its
/// `__control_name` field with `this.__control_name`. This way the
/// returned instance has all the method thunks (DrawLine etc.) AND
/// carries the source control's identity for canvas routing.
///
/// Stack trace:
/// ```text
///   PushThis
///   PushThisField "__control_name"     ; [this, name]
///   NewDotnet Graphics 0               ; [this, name, graphics]
///   ; need to swap graphics and name to do struct_set graphics.__control_name = name
///   ; ... but the DSL doesn't have swap.
/// ```
///
/// Workaround: stash `name` in a host-side intermediate via
/// `createGraphics`. The `vybe:gui::createGraphics` host fn returns a
/// pre-stamped Graphics-shaped Object. Then copy its `__control_name`
/// onto a fresh `NewDotnet Graphics` instance via struct_set.
///
/// Actually simpler: just emit `NewDotnet Graphics 0` then `SetField
/// __control_name` with `this.__control_name` on the stack ABOVE the
/// graphics instance. struct_set takes [obj, val] — so we need
/// graphics on the bottom, name on top. Build it as:
///
/// ```text
///   NewDotnet Graphics 0          ; [graphics]
///   PushThisField "__control_name" ; [graphics, name]
///   SetField "__control_name"      ; [name]   (struct_set leaves the val)
///   Drop                           ; []
///   NewDotnet Graphics 0           ; ... wait, we need to return the stamped graphics
/// ```
///
/// The struct_set leaves `val` on the stack, not `obj`. To return the
/// graphics with its stamped name, we need to dup it before
/// stamping:
///
/// ```text
///   NewDotnet Graphics 0          ; [graphics]
///   Dup                           ; [graphics, graphics]
///   PushThisField "__control_name" ; [graphics, graphics, name]
///   SetField "__control_name"      ; [graphics, name]
///   Drop                           ; [graphics]
///   Return                        ; returns graphics
/// ```
///
/// This works.
const CONTROL_CREATE_GRAPHICS: &[MethodOp] = &[
    // graphics = New Graphics()
    MethodOp::NewDotnet {
        class: "Graphics",
        argc: 0,
    },
    // Stamp graphics.__control_name = this.__control_name so subsequent
    // canvas calls route to this control's RecordingCanvas.
    MethodOp::Dup,
    MethodOp::PushThisField("__control_name"),
    MethodOp::SetField("__control_name"),
    MethodOp::Drop,
    MethodOp::Return,
];

/// Methods owned by `Control` (inherited by every concrete control,
/// including `Form`). Each entry maps to a host fn or, for compound
/// methods like `CreateGraphics`, to a [`MethodTarget::Body`] sequence.
///
/// - `Show`/`Hide`/`Focus`/`Refresh`/… → host fns in `vybe:gui` (no-ops
///   in non-display contexts, real implementations under a GUI backend)
/// - `CreateGraphics` → a Body that constructs a `Graphics` dotnet
///   instance and stamps `__control_name` from `this`. The returned
///   instance has all `Graphics` methods bound (via the standard dotnet
///   ctor) AND carries the source control's identity, so subsequent
///   `g.DrawLine(...)` calls route to the right `RecordingCanvas`.
///
/// Methods specific to `Form` (`Close`, `ShowDialog`, `Activate`,
/// `CenterToScreen`) live on the `Form` class so subclasses inherit them
/// only when inheriting from `Form`.
const CONTROL_METHODS: &[DotnetMethod] = &[
    DotnetMethod {
        name: "Show",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_show"),
    },
    DotnetMethod {
        name: "Hide",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_hide"),
    },
    DotnetMethod {
        name: "Focus",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_focus"),
    },
    DotnetMethod {
        name: "Refresh",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_refresh"),
    },
    DotnetMethod {
        name: "Invalidate",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_invalidate"),
    },
    DotnetMethod {
        name: "Update",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_update"),
    },
    DotnetMethod {
        name: "BringToFront",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_bring_to_front"),
    },
    DotnetMethod {
        name: "SendToBack",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_send_to_back"),
    },
    DotnetMethod {
        name: "Select",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_focus"),
    },
    DotnetMethod {
        name: "Dispose",
        arity: 1,
        target: MethodTarget::host("vybe:gui", "__ctrl_dispose"),
    },
    DotnetMethod {
        name: "SuspendLayout",
        arity: 1,
        target: MethodTarget::body(CONTROL_NOOP),
    },
    DotnetMethod {
        name: "ResumeLayout",
        arity: 2,
        target: MethodTarget::body(CONTROL_NOOP),
    },
    DotnetMethod {
        name: "PerformLayout",
        arity: 1,
        target: MethodTarget::body(CONTROL_NOOP),
    },
    DotnetMethod {
        name: "CreateGraphics",
        arity: 1,
        target: MethodTarget::body(CONTROL_CREATE_GRAPHICS),
    },
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
            properties: &["ActiveControl", "ParentForm", "AutoValidate"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
    ]
}
