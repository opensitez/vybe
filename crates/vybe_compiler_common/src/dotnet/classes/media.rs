//! Image / web media controls.

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "PictureBox",
            parent: Some("Control"),
            properties: &[
                "BorderStyle",
                "ErrorImage",
                "Image",
                "ImageLocation",
                "InitialImage",
                "SizeMode",
                "WaitOnLoad",
            ],
            widget_host_fn: Some("new_PictureBox"),
        },
        DotnetClass {
            name: "WebBrowser",
            parent: Some("Control"),
            properties: &[
                "AllowNavigation",
                "AllowWebBrowserDrop",
                "CanGoBack",
                "CanGoForward",
                "Document",
                "DocumentStream",
                "DocumentText",
                "DocumentTitle",
                "DocumentType",
                "EncryptionLevel",
                "IsBusy",
                "IsOffline",
                "IsWebBrowserContextMenuEnabled",
                "ObjectForScripting",
                "ReadyState",
                "ScriptErrorsSuppressed",
                "ScrollBarsEnabled",
                "StatusText",
                "Url",
                "Version",
                "WebBrowserShortcutsEnabled",
            ],
            widget_host_fn: Some("new_WebBrowser"),
        },
    ]
}
