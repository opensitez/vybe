//! .NET implicit import list.
//!
//! Owns the single piece of data that represents "what namespaces does every
//! .NET compiler implicitly recognise". Language compilers extend this with
//! language-specific additions (e.g. `microsoft.visualbasic` for VB,
//! `system.linq` for C#) before handing it to the resolver.
//!
//! This file is intentionally narrow — namespace-root recognition lives in
//! `namespaces.rs`, and namespace → host mapping lives in `host_map.rs`.

/// Return the default set of .NET namespace imports that every .NET compiler
/// should recognise. Returned as a Vec so callers can `.extend()` with extras.
pub fn default_interface_imports() -> Vec<String> {
    vec![
        "system".into(),
        "system.console".into(),
        "system.convert".into(),
        "system.math".into(),
        "system.string".into(),
        "system.array".into(),
        "system.environment".into(),
        // IO
        "system.io".into(),
        "system.io.file".into(),
        "system.io.path".into(),
        "system.io.directory".into(),
        // Collections
        "system.collections".into(),
        "system.collections.generic".into(),
        // Text
        "system.text".into(),
        "system.text.regularexpressions".into(),
        // Threading
        "system.threading".into(),
        "system.threading.thread".into(),
        "system.threading.tasks".into(),
        // Diagnostics
        "system.diagnostics".into(),
        "system.diagnostics.process".into(),
        "system.diagnostics.stopwatch".into(),
        "system.diagnostics.debug".into(),
        "system.diagnostics.trace".into(),
        // Drawing
        "system.drawing".into(),
        // WinForms
        "system.windows.forms".into(),
        // Net
        "system.net".into(),
        "system.net.sockets".into(),
        // Data
        "system.data".into(),
        "system.data.sqlclient".into(),
        "system.data.oledb".into(),
        // Security
        "system.security.cryptography".into(),
        // XML
        "system.xml.linq".into(),
        // LINQ
        "system.linq".into(),
        // WinForms bare names (for Application.Run, Application.Exit)
        "application".into(),
    ]
}
