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
            properties: &["BorderStyle"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_Panel"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "GroupBox",
            parent: Some("Control"),
            properties: &["FlatStyle", "UseCompatibleTextRendering"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_GroupBox"),
            widget_host_module: "vybe:gui",
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
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_TabControl"),
            widget_host_module: "vybe:gui",
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
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_TabPage"),
            widget_host_module: "vybe:gui",
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
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_SplitContainer"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "FlowLayoutPanel",
            parent: Some("Panel"),
            properties: &["FlowDirection", "WrapContents"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_FlowLayoutPanel"),
            widget_host_module: "vybe:gui",
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
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_TableLayoutPanel"),
            widget_host_module: "vybe:gui",
        },
        // ── Declared at last: the designer knew them, the descriptor did not ──
        //
        // `ControlType` (`winforms/control.rs`) has carried `HScrollBar`,
        // `VScrollBar` and `BindingNavigator` all along, so a `.Designer.vb`
        // could name them — but with no class here they had no tree `Type`, so
        // `control_element_for_type` answered None and **every property write
        // on them was dropped**. In `examples/vb/allcontrols` that is exactly
        // what happened: `Me.hsb1.Location = New Point(230, 850)` went to a
        // plain object property and all three sat at the origin, stacked on
        // the strips, with their designer names never applied either.
        //
        // `widget_host_fn: None` on purpose — the ELEMENT materialises them
        // (`is_element_mapped`), which is the direction the whole conversion
        // runs in. `vybe_widgets` already has all three kinds and their default
        // sizes; only the declaration was missing.
        DotnetClass {
            name: "HScrollBar",
            parent: Some("Control"),
            properties: &[
                "LargeChange",
                "Maximum",
                "Minimum",
                "SmallChange",
                "Value",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "VScrollBar",
            parent: Some("Control"),
            properties: &[
                "LargeChange",
                "Maximum",
                "Minimum",
                "SmallChange",
                "Value",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        // A BindingNavigator IS a ToolStrip — that is its real .NET parent, and
        // saying so is what gives it the strip's `Items` surface for free.
        DotnetClass {
            name: "BindingNavigator",
            parent: Some("ToolStrip"),
            properties: &["BindingSource", "CountItem", "PositionItem"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        // The bare drag-bar `Splitter` — plib's `TSplitter`, and NOT
        // `SplitContainer`, which is the two-panel container above.
        DotnetClass {
            name: "Splitter",
            parent: Some("Control"),
            properties: &["BorderStyle", "MinExtra", "MinSize", "SplitPosition"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        // A `UserControl` is a composite the program itself fills, so it is a
        // plain container — `<section>`, which is a real element and already
        // establishes a containing block.
        DotnetClass {
            name: "UserControl",
            parent: Some("Panel"),
            properties: &["AutoScaleMode", "AutoValidate", "BorderStyle"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        // `DomainUpDown` is `NumericUpDown`'s text-list twin.
        DotnetClass {
            name: "DomainUpDown",
            parent: Some("Control"),
            properties: &[
                "Items",
                "ReadOnly",
                "SelectedIndex",
                "SelectedItem",
                "Sorted",
                "Wrap",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
    ]
}
