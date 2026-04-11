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
            widget_host_fn: Some("new_ToolStrip"),
        },
        DotnetClass {
            name: "MenuStrip",
            parent: Some("ToolStrip"),
            properties: &[
                "MdiWindowListItem",
                "ShowItemToolTips",
                "Stretch",
            ],
            widget_host_fn: Some("new_MenuStrip"),
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
            widget_host_fn: Some("new_StatusStrip"),
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
            widget_host_fn: Some("new_ContextMenuStrip"),
        },
    ]
}
