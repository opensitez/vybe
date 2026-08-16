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

/// `Control.CreateGraphics()` body — **`element.getContext("2d")`**,
/// HTML §4.12.5.
///
/// ```text
///   PushThis                      ; [control]      — the control IS an element
///   PushConstStr "2d"             ; [control, "2d"]
///   CallHost web:canvas getContext 2  ; [graphics]
///   Dup / CallHost save 1 / Drop  ; [graphics]     — the clip baseline
///   Dup                           ; [graphics, graphics]
///   PushConstStr "Graphics"       ; [graphics, graphics, "Graphics"]
///   SetField "__type"             ; [graphics, "Graphics"]  (struct_set leaves the val)
///   Drop                          ; [graphics]
///   Return
/// ```
///
/// The `save` is the clip baseline `SetClip`/`ResetClip` pop back to — a
/// canvas clip has no inverse, so the region is undone by returning to a state
/// saved before it was applied. Pushing it once here is what makes those two
/// unable to underflow the state stack. See `drawing.rs::GRAPHICS_SET_CLIP`.
///
/// A control is built by `emit_control_element` with
/// `document.createElement`, so the receiver already carries `__node` and a
/// context can be asked for directly. That is the only form of the call a
/// real browser engine can answer: there is no control name to resolve on the
/// other side of the seam.
///
/// The `"2d"` is not decoration: `getContext` answers `null` for a context
/// type it does not support, per spec, and an absent argument reads as `""`.
/// A null context still accepts every drawing call and paints nothing, so the
/// argument is what stands between a working surface and a silent one.
///
/// `__type` is re-stamped because guest code downcasts on it; the handle
/// introduces itself as `CanvasRenderingContext2D`, which is what it is and
/// not what the .NET object wrapping it is called. `Graphics` methods resolve
/// through the class descriptor by static type, not off the instance, so the
/// returned handle needs no bound thunks — only its identity.
///
/// Kept in step with `dispatch.rs::emit_control_create_graphics`, which is
/// the same lowering in raw bytecode for the descriptor route
/// (`dotnet.control_create_graphics`). Both must agree or a control's
/// `Graphics` depends on which route compiled it.
const CONTROL_CREATE_GRAPHICS: &[MethodOp] = &[
    MethodOp::PushThis,
    MethodOp::PushConstStr("2d"),
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "getContext",
        argc: 2,
    },
    MethodOp::Dup,
    MethodOp::CallHost {
        module: "web:canvas",
        fn_name: "save",
        argc: 1,
    },
    MethodOp::Drop,
    // `SetField` lowers to `STRUCT_SET`, which consumes the object AND the
    // value — the `Dup` above is what leaves the context behind, so there is
    // nothing to `Drop`. A `Drop` here empties the stack and the method
    // returns `null`.
    MethodOp::Dup,
    MethodOp::PushConstStr("Graphics"),
    MethodOp::SetField("__type"),
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
