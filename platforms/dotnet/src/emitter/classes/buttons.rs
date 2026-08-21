//! `ButtonBase → {Button, CheckBox, RadioButton}`.
//!
//! `ButtonBase` is the abstract base for all button-like controls. Real
//! .NET puts `FlatStyle`, `Image*`, `TextAlign`, `TextImageRelation` here
//! so they're shared across `Button` and the check/radio variants.
//!
//! `Button` adds `DialogResult`. `CheckBox` and `RadioButton` add the
//! `Checked` / `CheckState` family. All three are concrete leaves backed by an
//! ELEMENT, not a host factory — `<button>` and `<input type="checkbox">` /
//! `<input type="radio">`, declared in `tree_register::html_element_for_control`.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "ButtonBase",
            parent: Some("Control"),
            properties: &[
                "FlatStyle",
                "FlatAppearance",
                "Image",
                "ImageAlign",
                "ImageIndex",
                "ImageKey",
                "ImageList",
                "TextAlign",
                "TextImageRelation",
                "UseCompatibleTextRendering",
                "UseMnemonic",
                "UseVisualStyleBackColor",
                "AutoEllipsis",
                "IsDefault",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "Button",
            parent: Some("ButtonBase"),
            properties: &["DialogResult", "AutoSizeMode"],
            methods: &[],
            ctor_arity: 0,
            // `<button>` — materialized by the element mapping in
            // `tree_register::html_element_for_control`. A `widget_host_fn`
            // here would win over that mapping and pin the control to a host
            // factory; see `winforms::component_classes`.
            widget_host_fn: None,        },
        DotnetClass {
            name: "CheckBox",
            parent: Some("ButtonBase"),
            properties: &[
                "Appearance",
                "AutoCheck",
                "CheckAlign",
                "Checked",
                "CheckState",
                "ThreeState",
            ],
            methods: &[],
            ctor_arity: 0,
            // `<input type="checkbox">`
            widget_host_fn: None,        },
        DotnetClass {
            name: "RadioButton",
            parent: Some("ButtonBase"),
            properties: &["Appearance", "AutoCheck", "CheckAlign", "Checked"],
            methods: &[],
            ctor_arity: 0,
            // `<input type="radio">`
            widget_host_fn: None,        },
    ]
}
