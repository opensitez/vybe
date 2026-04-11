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
            properties: &[
                "Tag",
            ],
            widget_host_fn: None,
        },
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
            widget_host_fn: None,
        },
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
            widget_host_fn: Some("new_OpenFileDialog"),
        },
        DotnetClass {
            name: "SaveFileDialog",
            parent: Some("FileDialog"),
            properties: &[
                "CreatePrompt",
                "OverwritePrompt",
            ],
            widget_host_fn: Some("new_SaveFileDialog"),
        },
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
            widget_host_fn: Some("new_FontDialog"),
        },
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
            widget_host_fn: Some("new_ColorDialog"),
        },
        DotnetClass {
            name: "FolderBrowserDialog",
            parent: Some("CommonDialog"),
            properties: &[
                "Description",
                "RootFolder",
                "SelectedPath",
                "ShowNewFolderButton",
            ],
            widget_host_fn: Some("new_FolderBrowserDialog"),
        },
    ]
}
