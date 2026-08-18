//! `TextBoxBase → {TextBox, RichTextBox, MaskedTextBox}`.
//!
//! `TextBoxBase` owns the shared editing surface (selection, scrollbars,
//! readonly, max length). The concrete subclasses add their own quirks.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "TextBoxBase",
            parent: Some("Control"),
            properties: &[
                "AcceptsTab",
                "AutoSize",
                "BorderStyle",
                "CanUndo",
                "HideSelection",
                "Lines",
                "MaxLength",
                "Modified",
                "Multiline",
                "ReadOnly",
                "ScrollBars",
                "SelectedText",
                "SelectionLength",
                "SelectionStart",
                "ShortcutsEnabled",
                "TextLength",
                "WordWrap",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "TextBox",
            parent: Some("TextBoxBase"),
            properties: &[
                "AcceptsReturn",
                "AutoCompleteCustomSource",
                "AutoCompleteMode",
                "AutoCompleteSource",
                "CharacterCasing",
                "PasswordChar",
                "PlaceholderText",
                "TextAlign",
                "UseSystemPasswordChar",
            ],
            methods: &[],
            ctor_arity: 0,
            // `<input type="text">` — created by the element mapping.
            widget_host_fn: None,        },
        DotnetClass {
            name: "RichTextBox",
            parent: Some("TextBoxBase"),
            properties: &[
                "AutoWordSelection",
                "BulletIndent",
                "DetectUrls",
                "EnableAutoDragDrop",
                "RightMargin",
                "Rtf",
                "ScrollBars",
                "SelectionAlignment",
                "SelectionBackColor",
                "SelectionBullet",
                "SelectionColor",
                "SelectionFont",
                "SelectionIndent",
                "ShowSelectionMargin",
                "ZoomFactor",
            ],
            methods: &[],
            ctor_arity: 0,
            // `<textarea>` — the multiline text surface.
            widget_host_fn: None,        },
        DotnetClass {
            name: "MaskedTextBox",
            parent: Some("TextBoxBase"),
            properties: &[
                "AllowPromptAsInput",
                "AsciiOnly",
                "BeepOnError",
                "CutCopyMaskFormat",
                "Culture",
                "HidePromptOnLeave",
                "InsertKeyMode",
                "Mask",
                "MaskCompleted",
                "MaskFull",
                "PromptChar",
                "RejectInputOnFirstFailure",
                "ResetOnPrompt",
                "ResetOnSpace",
                "SkipLiterals",
                "TextAlign",
                "TextMaskFormat",
                "ValidatingType",
            ],
            methods: &[],
            ctor_arity: 0,
            // `<input type="text">`
            widget_host_fn: None,        },
    ]
}
