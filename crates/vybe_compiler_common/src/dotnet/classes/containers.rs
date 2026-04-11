//! Container controls: `Panel`, `GroupBox`, `TabControl`, `TabPage`,
//! `SplitContainer`, `FlowLayoutPanel`, `TableLayoutPanel`.
//!
//! `Panel` inherits from `ScrollableControl` (it has its own scrollbars).
//! `TabPage`, `FlowLayoutPanel`, `TableLayoutPanel`, `SplitterPanel`
//! inherit from `Panel`. `GroupBox` is a borderless container with a
//! caption — inherits straight from `Control`. `TabControl` inherits from
//! `Control` and owns the tab strip.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "Panel",
            parent: Some("ScrollableControl"),
            properties: &[
                "BorderStyle",
            ],
            widget_host_fn: Some("new_Panel"),
        },
        DotnetClass {
            name: "GroupBox",
            parent: Some("Control"),
            properties: &[
                "FlatStyle",
                "UseCompatibleTextRendering",
            ],
            widget_host_fn: Some("new_GroupBox"),
        },
        DotnetClass {
            name: "TabControl",
            parent: Some("Control"),
            properties: &[
                "Alignment",
                "Appearance",
                "DrawMode",
                "HotTrack",
                "ImageList",
                "ItemSize",
                "Multiline",
                "Padding",
                "RowCount",
                "SelectedIndex",
                "SelectedTab",
                "ShowToolTips",
                "SizeMode",
                "TabCount",
                "TabPages",
            ],
            widget_host_fn: Some("new_TabControl"),
        },
        DotnetClass {
            name: "TabPage",
            parent: Some("Panel"),
            properties: &[
                "ImageIndex",
                "ImageKey",
                "ToolTipText",
                "UseVisualStyleBackColor",
            ],
            widget_host_fn: Some("new_TabPage"),
        },
        DotnetClass {
            name: "SplitContainer",
            parent: Some("ContainerControl"),
            properties: &[
                "BorderStyle",
                "FixedPanel",
                "IsSplitterFixed",
                "Orientation",
                "Panel1",
                "Panel1Collapsed",
                "Panel1MinSize",
                "Panel2",
                "Panel2Collapsed",
                "Panel2MinSize",
                "SplitterDistance",
                "SplitterIncrement",
                "SplitterRectangle",
                "SplitterWidth",
            ],
            widget_host_fn: Some("new_SplitContainer"),
        },
        DotnetClass {
            name: "FlowLayoutPanel",
            parent: Some("Panel"),
            properties: &[
                "FlowDirection",
                "WrapContents",
            ],
            widget_host_fn: Some("new_FlowLayoutPanel"),
        },
        DotnetClass {
            name: "TableLayoutPanel",
            parent: Some("Panel"),
            properties: &[
                "CellBorderStyle",
                "ColumnCount",
                "ColumnStyles",
                "Controls",
                "GrowStyle",
                "RowCount",
                "RowStyles",
            ],
            widget_host_fn: Some("new_TableLayoutPanel"),
        },
    ]
}
