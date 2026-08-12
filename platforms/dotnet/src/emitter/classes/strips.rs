//! Menu / tool / status / context strips, and the ITEMS that go on them.
//!
//! In real .NET they all inherit from `ToolStrip` (`MenuStrip`, `StatusStrip`,
//! `ContextMenuStrip` are subclasses). We model the same hierarchy.
//!
//! The items sit on a SEPARATE .NET hierarchy — `ToolStripItem` derives from
//! `Component`, not from `Control` — and that is modelled truthfully here,
//! because `mi is Control` must answer False the way .NET answers it. What
//! makes them render anyway is that they are element-mapped
//! (`html_element_for_control`), which is the question the document actually
//! asks; see `element_backed_control` in the registrar.

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
            properties: &["MdiWindowListItem", "ShowItemToolTips", "Stretch"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: Some("new_MenuStrip"),
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "StatusStrip",
            parent: Some("ToolStrip"),
            properties: &["LayoutStyle", "ShowItemToolTips", "SizingGrip", "Stretch"],
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
        // ── The items ──────────────────────────────────────────────────────
        // `ToolStripItem` is .NET's shared base for everything that sits ON a
        // strip. It is a `Component`, NOT a `Control`, which is why it carries
        // its own `Text`/`Enabled`/`Visible` rather than inheriting them.
        DotnetClass {
            name: "ToolStripItem",
            parent: Some("Component"),
            properties: &[
                "AccessibleName",
                "Alignment",
                "AutoSize",
                "Available",
                "BackColor",
                "DisplayStyle",
                "Enabled",
                "Font",
                "ForeColor",
                "Image",
                "ImageAlign",
                "Name",
                "Owner",
                "Padding",
                "Tag",
                "Text",
                "TextAlign",
                "ToolTipText",
                "Visible",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        // `ToolStripDropDownItem` is what makes an item OPEN something —
        // declared so `DropDownItems` resolves, even though drop-downs do not
        // open yet.
        DotnetClass {
            name: "ToolStripDropDownItem",
            parent: Some("ToolStripItem"),
            properties: &["DropDown", "DropDownItems", "HasDropDownItems"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "ToolStripMenuItem",
            parent: Some("ToolStripDropDownItem"),
            properties: &[
                "CheckOnClick",
                "CheckState",
                "Checked",
                "ShortcutKeyDisplayString",
                "ShortcutKeys",
                "ShowShortcutKeys",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "ToolStripButton",
            parent: Some("ToolStripItem"),
            properties: &["CheckOnClick", "CheckState", "Checked"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "ToolStripLabel",
            parent: Some("ToolStripItem"),
            properties: &["IsLink", "LinkColor", "LinkVisited"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "ToolStripStatusLabel",
            parent: Some("ToolStripLabel"),
            properties: &["BorderSides", "BorderStyle", "Spring"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "ToolStripSeparator",
            parent: Some("ToolStripItem"),
            properties: &[],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
    ]
}
