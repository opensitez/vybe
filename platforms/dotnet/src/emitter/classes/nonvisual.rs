//! Non-visual / lifecycle controls.
//!
//! These don't render but live on the form designer surface as components.
//! In real .NET they inherit from `Component` (not `Control`) because they
//! have no UI. They get their own `__set_<prop>` chain even though they
//! never render — user code still does `bs.DataSource = ...`.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "Timer",
            parent: Some("Component"),
            properties: &["Enabled", "Interval", "Tag"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_Timer"),
            widget_host_module: "vybe:gui",
        },
        // `BindingSource` is NOT declared here. It is a cursor over data, not a
        // control: it has no element, nothing paints it, and every member is a
        // position or a list. Declared in this table it got a `vybe:gui`
        // backing object and property accessors keyed by CONTROL NAME, so
        // `bs.Position` read back `""` and `bs.MoveFirst()` was `undefined`
        // (`methods: &[]`). It lives in
        // `core/component_classes_data_drawing.rs` next to `DataTable` and
        // `DataRow`, with real fields and real emits — see
        // `core/bindingsource_adapter.rs`.
        DotnetClass {
            name: "ImageList",
            parent: Some("Component"),
            properties: &[
                "ColorDepth",
                "Container",
                "Handle",
                "HandleCreated",
                "Images",
                "ImageSize",
                "ImageStream",
                "Tag",
                "TransparentColor",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_ImageList"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "ToolTip",
            parent: Some("Component"),
            properties: &[
                "Active",
                "AutomaticDelay",
                "AutoPopDelay",
                "BackColor",
                "ForeColor",
                "InitialDelay",
                "IsBalloon",
                "OwnerDraw",
                "ReshowDelay",
                "ShowAlways",
                "StripAmpersands",
                "Tag",
                "ToolTipIcon",
                "ToolTipTitle",
                "UseAnimation",
                "UseFading",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_ToolTip"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "NotifyIcon",
            parent: Some("Component"),
            properties: &[
                "BalloonTipIcon",
                "BalloonTipText",
                "BalloonTipTitle",
                "ContextMenuStrip",
                "Icon",
                "Tag",
                "Text",
                "Visible",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_NotifyIcon"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "ErrorProvider",
            parent: Some("Component"),
            properties: &[
                "BlinkRate",
                "BlinkStyle",
                "ContainerControl",
                "DataMember",
                "DataSource",
                "Icon",
                "RightToLeft",
                "Tag",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_ErrorProvider"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "HelpProvider",
            parent: Some("Component"),
            properties: &["HelpNamespace", "Tag"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_HelpProvider"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "BackgroundWorker",
            parent: Some("Component"),
            properties: &[
                "CancellationPending",
                "IsBusy",
                "WorkerReportsProgress",
                "WorkerSupportsCancellation",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_BackgroundWorker"),
            widget_host_module: "vybe:gui",
        },
    ]
}
