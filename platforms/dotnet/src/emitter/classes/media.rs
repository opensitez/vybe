//! Image / web media controls.
//!
//! `PictureBox` is the WinForms paintable image surface. Real .NET
//! supports both static images (`pb.Image = New Bitmap(...)`) and ad-hoc
//! drawing via `pb.CreateGraphics()`. Both flow through the same
//! underlying surface — and that surface is exactly what the
//! `vybe_widgets::Canvas` widget provides.
//!
//! So `PictureBox` is wired to `vybe:gui::new_Canvas` (which constructs
//! a `vybe_widgets::Canvas` widget). User code that does
//! `pb.CreateGraphics()` gets a `Graphics` handle pointing at the
//! Canvas widget's own `RecordingCanvas` — drawings persist between
//! frames, the form's render loop replays them onto the live pixmap,
//! and the user sees the painting on screen.

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
            // Backed by the bare Canvas widget — see
            // `vybe_widgets::canvas_widget::Canvas`. The user gets the
            // full `Control` method chain (`Show`, `Hide`,
            // `CreateGraphics`, …) plus an actual paintable surface.
            widget_host_fn: Some("new_Canvas"),
            widget_host_module: "vybe:gui" },
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
            widget_host_fn: Some("new_WebBrowser"),
            widget_host_module: "vybe:gui" },
    ]
}
