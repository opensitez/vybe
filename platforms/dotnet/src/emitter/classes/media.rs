//! Image / web media controls.
//!
//! `PictureBox` is the WinForms paintable image surface. Real .NET
//! supports both static images (`pb.Image = New Bitmap(...)`) and ad-hoc
//! drawing via `pb.CreateGraphics()`. Both flow through the same
//! underlying surface — and on the web that surface is `<canvas>`.
//!
//! So `PictureBox` IS a `<canvas>` element: it is created by
//! `document.createElement("canvas")` through the element mapping in
//! `tree_register::html_element_for_control`, and `pb.CreateGraphics()` is
//! `getContext("2d")` on that same element. No `vybe:gui` factory stands in
//! between, which is what makes the whole path answerable by a real browser.

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
            methods: &[],
            ctor_arity: 0,
            // No factory: `picturebox` maps to `canvas`, and the element
            // mapping is what materializes it. `component_classes` checks
            // `widget_host_fn` FIRST, so leaving one here would keep the
            // control on `vybe:gui::new_Canvas` and the element mapping would
            // never be reached — only a `<canvas>` tag owns a drawing surface.
            widget_host_fn: None,        },
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
            methods: &[],
            ctor_arity: 0,
            // `<iframe>` — see `html_element_for_control`. Renders as a plain
            // box until `vybe_widgets` grows a `webbrowser` kind.
            widget_host_fn: None,        },
    ]
}
