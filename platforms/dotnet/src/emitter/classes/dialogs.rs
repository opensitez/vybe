//! WinForms common dialogs.
//!
//! `OpenFileDialog`, `SaveFileDialog`, `FontDialog`, `ColorDialog`,
//! `FolderBrowserDialog`. They all inherit from `CommonDialog` which
//! inherits from `Component`.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "CommonDialog",
            parent: Some("Component"),
            properties: &["Tag"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "FileDialog",
            parent: Some("CommonDialog"),
            properties: &[
                "AddExtension",
                "AutoUpgradeEnabled",
                "CheckFileExists",
                "CheckPathExists",
                "CustomPlaces",
                "DefaultExt",
                "DereferenceLinks",
                "FileName",
                "FileNames",
                "Filter",
                "FilterIndex",
                "InitialDirectory",
                "RestoreDirectory",
                "ShowHelp",
                "SupportMultiDottedExtensions",
                "Title",
                "ValidateNames",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "OpenFileDialog",
            parent: Some("FileDialog"),
            properties: &[
                "Multiselect",
                "ReadOnlyChecked",
                "SafeFileName",
                "SafeFileNames",
                "ShowReadOnly",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "SaveFileDialog",
            parent: Some("FileDialog"),
            properties: &["CreatePrompt", "OverwritePrompt"],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "FontDialog",
            parent: Some("CommonDialog"),
            properties: &[
                "AllowScriptChange",
                "AllowSimulations",
                "AllowVectorFonts",
                "AllowVerticalFonts",
                "Color",
                "FixedPitchOnly",
                "Font",
                "FontMustExist",
                "MaxSize",
                "MinSize",
                "ScriptsOnly",
                "ShowApply",
                "ShowColor",
                "ShowEffects",
                "ShowHelp",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "ColorDialog",
            parent: Some("CommonDialog"),
            properties: &[
                "AllowFullOpen",
                "AnyColor",
                "Color",
                "CustomColors",
                "FullOpen",
                "ShowHelp",
                "SolidColorOnly",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
        DotnetClass {
            name: "FolderBrowserDialog",
            parent: Some("CommonDialog"),
            properties: &[
                "Description",
                "RootFolder",
                "SelectedPath",
                "ShowNewFolderButton",
            ],
            methods: &[],
            ctor_arity: 0,
            widget_host_fn: None,        },
    ]
}
