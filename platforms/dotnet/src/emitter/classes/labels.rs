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
            methods: &[],
            ctor_arity: 0,
            // `<label>` — created by the element mapping, not a factory.
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
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
            methods: &[],
            ctor_arity: 0,
            // `<a>`
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
    ]
}
