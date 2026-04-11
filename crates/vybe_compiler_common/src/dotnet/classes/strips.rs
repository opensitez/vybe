//! Menu / tool / status / context strips.
//!
//! In real .NET they all inherit from `ToolStrip` (`MenuStrip`, `StatusStrip`,
//! `ContextMenuStrip` are subclasses). We model the same hierarchy.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "ToolStrip",
            parent: Some("ScrollableControl"),
            properties: &[
                "AllowDrop",
                "AllowItemReorder",
                "AllowMerge",
                "AutoSize",
                "BackgroundImage",
                "BackgroundImageLayout",
                "CanOverflow",
                "DefaultDropDownDirection",
                "Dock",
                "GripDisplayStyle",
                "GripMargin",
                "GripStyle",
                "ImageList",
                "ImageScalingSize",
                "Items",
                "LayoutEngine",
                "LayoutSettings",
                "LayoutStyle",
                "Renderer",
                "RenderMode",
                "ShowItemToolTips",
                "Stretch",
                "TabStop",
                "TextDirection",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_ToolStrip"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "MenuStrip",
            parent: Some("ToolStrip"),
            properties: &[
                "MdiWindowListItem",
                "ShowItemToolTips",
                "Stretch",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_MenuStrip"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "StatusStrip",
            parent: Some("ToolStrip"),
            properties: &[
                "LayoutStyle",
                "ShowItemToolTips",
                "SizingGrip",
                "Stretch",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_StatusStrip"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "ContextMenuStrip",
            parent: Some("ToolStrip"),
            properties: &[
                "AutoClose",
                "AutoSize",
                "DropShadowEnabled",
                "OwnerItem",
                "SourceControl",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_ContextMenuStrip"),
            widget_host_module: "vybe:gui",
        },
    ]
}
