//! Non-visual / lifecycle controls.
//!
//! These don't render but live on the form designer surface as components.
//! In real .NET they inherit from `Component` (not `Control`) because they
//! have no UI. They get their own `__set_<prop>` chain even though they
//! never render — user code still does `bs.DataSource = ...`.

use super::{DotnetClass, DotnetMethod, MethodTarget};

const TOOLTIP_METHODS: &[DotnetMethod] = &[
    DotnetMethod {
        name: "SetToolTip",
        arity: 3,
        target: MethodTarget::Common {
            emit: "dotnet.tooltip_set",
        },
    },
    DotnetMethod {
        name: "GetToolTip",
        arity: 2,
        target: MethodTarget::Common {
            emit: "dotnet.tooltip_get",
        },
    },
    DotnetMethod {
        name: "RemoveAll",
        arity: 1,
        target: MethodTarget::Common {
            emit: "dotnet.tooltip_remove_all",
        },
    },
];

const HELP_PROVIDER_METHODS: &[DotnetMethod] = &[
    DotnetMethod { name: "SetHelpString", arity: 3, target: MethodTarget::Common { emit: "dotnet.help_set.string" } },
    DotnetMethod { name: "GetHelpString", arity: 2, target: MethodTarget::Common { emit: "dotnet.help_get.string" } },
    DotnetMethod { name: "SetHelpKeyword", arity: 3, target: MethodTarget::Common { emit: "dotnet.help_set.keyword" } },
    DotnetMethod { name: "GetHelpKeyword", arity: 2, target: MethodTarget::Common { emit: "dotnet.help_get.keyword" } },
    DotnetMethod { name: "SetHelpNavigator", arity: 3, target: MethodTarget::Common { emit: "dotnet.help_set.navigator" } },
    DotnetMethod { name: "GetHelpNavigator", arity: 2, target: MethodTarget::Common { emit: "dotnet.help_get.navigator" } },
    DotnetMethod { name: "SetShowHelp", arity: 3, target: MethodTarget::Common { emit: "dotnet.help_set.show" } },
    DotnetMethod { name: "GetShowHelp", arity: 2, target: MethodTarget::Common { emit: "dotnet.help_get.show" } },
];

const ERROR_PROVIDER_METHODS: &[DotnetMethod] = &[
    DotnetMethod { name: "SetError", arity: 3, target: MethodTarget::Common { emit: "dotnet.error_provider_set" } },
    DotnetMethod { name: "GetError", arity: 2, target: MethodTarget::Common { emit: "dotnet.error_provider_get" } },
    DotnetMethod { name: "Clear", arity: 1, target: MethodTarget::Common { emit: "dotnet.error_provider_clear" } },
];

const WINFORMS_TIMER_METHODS: &[DotnetMethod] = &[
    DotnetMethod { name: "Start", arity: 1, target: MethodTarget::Common { emit: "dotnet.winforms_timer_start" } },
    DotnetMethod { name: "Stop", arity: 1, target: MethodTarget::Common { emit: "dotnet.winforms_timer_stop" } },
    DotnetMethod { name: "Dispose", arity: 1, target: MethodTarget::Common { emit: "dotnet.winforms_timer_stop" } },
];

const BACKGROUND_WORKER_METHODS: &[DotnetMethod] = &[
    DotnetMethod { name: "RunWorkerAsync", arity: 1, target: MethodTarget::Common { emit: "dotnet.background_worker_run" } },
    DotnetMethod { name: "RunWorkerAsync", arity: 2, target: MethodTarget::Common { emit: "dotnet.background_worker_run" } },
    DotnetMethod { name: "CancelAsync", arity: 1, target: MethodTarget::Common { emit: "dotnet.background_worker_cancel" } },
    DotnetMethod { name: "ReportProgress", arity: 2, target: MethodTarget::Common { emit: "dotnet.background_worker_progress" } },
    DotnetMethod { name: "ReportProgress", arity: 3, target: MethodTarget::Common { emit: "dotnet.background_worker_progress" } },
];

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "Timer",
            parent: Some("Component"),
            properties: &["Enabled", "Interval", "Tag"],
            methods: WINFORMS_TIMER_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,
        },
        // `BindingSource` is NOT declared here. It is a cursor over data, not a
        // control: it has no element, nothing paints it, and every member is a
        // position or a list. Declared in this table it got a host
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
            widget_host_fn: None,
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
            methods: TOOLTIP_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,
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
            widget_host_fn: None,
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
            methods: ERROR_PROVIDER_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,
        },
        DotnetClass {
            name: "HelpProvider",
            parent: Some("Component"),
            properties: &["HelpNamespace", "Tag"],
            methods: HELP_PROVIDER_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,
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
            methods: BACKGROUND_WORKER_METHODS,
            ctor_arity: 0,
            widget_host_fn: None,
        },
    ]
}
