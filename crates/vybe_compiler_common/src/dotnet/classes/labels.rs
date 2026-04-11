//! `Label` and `LinkLabel`.
//!
//! `Label` inherits straight from `Control`. `LinkLabel` inherits from
//! `Label` (real .NET — `LinkLabel : Label : Control`) and adds the link
//! tracking surface.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "Label",
            parent: Some("Control"),
            properties: &[
                "AutoEllipsis",
                "AutoSize",
                "BorderStyle",
                "FlatStyle",
                "Image",
                "ImageAlign",
                "ImageIndex",
                "ImageKey",
                "ImageList",
                "PreferredHeight",
                "PreferredWidth",
                "TextAlign",
                "UseCompatibleTextRendering",
                "UseMnemonic",
            ],
            widget_host_fn: Some("new_Label"),
        },
        DotnetClass {
            name: "LinkLabel",
            parent: Some("Label"),
            properties: &[
                "ActiveLinkColor",
                "DisabledLinkColor",
                "LinkArea",
                "LinkBehavior",
                "LinkColor",
                "Links",
                "LinkVisited",
                "OverrideCursor",
                "VisitedLinkColor",
            ],
            widget_host_fn: Some("new_LinkLabel"),
        },
    ]
}
