//! `ButtonBase → {Button, CheckBox, RadioButton}`.
//!
//! `ButtonBase` is the abstract base for all button-like controls. Real
//! .NET puts `FlatStyle`, `Image*`, `TextAlign`, `TextImageRelation` here
//! so they're shared across `Button` and the check/radio variants.
//!
//! `Button` adds `DialogResult`. `CheckBox` and `RadioButton` add the
//! `Checked` / `CheckState` family. All three are concrete leaves backed
//! by the matching `vybe:gui::new_<Type>` host fn.

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
            widget_host_fn: None,
        },
        DotnetClass {
            name: "Button",
            parent: Some("ButtonBase"),
            properties: &[
                "DialogResult",
                "AutoSizeMode",
            ],
            widget_host_fn: Some("new_Button"),
        },
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
            widget_host_fn: Some("new_CheckBox"),
        },
        DotnetClass {
            name: "RadioButton",
            parent: Some("ButtonBase"),
            properties: &[
                "Appearance",
                "AutoCheck",
                "CheckAlign",
                "Checked",
            ],
            widget_host_fn: Some("new_RadioButton"),
        },
    ]
}
