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
            properties: &[
                "Enabled",
                "Interval",
                "Tag",
            ],
            widget_host_fn: Some("new_Timer"),
        },
        DotnetClass {
            name: "BindingSource",
            parent: Some("Component"),
            properties: &[
                "AllowNew",
                "Count",
                "CurrencyManager",
                "Current",
                "DataMember",
                "DataSource",
                "Filter",
                "IsBindingSuspended",
                "IsFixedSize",
                "IsReadOnly",
                "IsSorted",
                "IsSynchronized",
                "Item",
                "List",
                "Position",
                "RaiseListChangedEvents",
                "Sort",
                "SortDescriptions",
                "SortDirection",
                "SortProperty",
                "SupportsAdvancedSorting",
                "SupportsChangeNotification",
                "SupportsFiltering",
                "SupportsSearching",
                "SupportsSorting",
                "SyncRoot",
            ],
            widget_host_fn: Some("new_BindingSource"),
        },
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
            widget_host_fn: Some("new_ImageList"),
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
            widget_host_fn: Some("new_ToolTip"),
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
            widget_host_fn: Some("new_NotifyIcon"),
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
            widget_host_fn: Some("new_ErrorProvider"),
        },
        DotnetClass {
            name: "HelpProvider",
            parent: Some("Component"),
            properties: &[
                "HelpNamespace",
                "Tag",
            ],
            widget_host_fn: Some("new_HelpProvider"),
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
            widget_host_fn: Some("new_BackgroundWorker"),
        },
    ]
}
