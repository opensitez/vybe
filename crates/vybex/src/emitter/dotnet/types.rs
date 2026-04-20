//! .NET BCL type tables, predicates, and name-shape helpers.
//!
//! This is the static reference data the rest of the .NET frontend leans on:
//! - `known_types()` — bare type name → host constructor mapping for `New X()`
//! - `is_noop_method` — WinForms layout/lifecycle methods that compile to null
//! - `is_known_constant` — .NET property-like constants (Math.PI, etc.) that
//!   shouldn't be invoked even when args are empty
//! - `capitalize_control_name` / `capitalize_data_type` — name shape helpers
//!   used by callers that need PascalCase forms

use std::collections::HashMap;

/// WinForms layout/lifecycle methods that are always no-ops at runtime.
pub fn is_noop_method(name: &str) -> bool {
    matches!(name,
        "suspendlayout" | "resumelayout" | "performlayout" |
        "refresh" | "invalidate" | "update" | "begininit" | "endinit" |
        "dispose" | "select" | "focus" | "bringtofront" | "sendtoback" |
        "createcontrol" | "show" | "hide"
    )
}

/// .NET property-like constants that should NOT be called even when args are empty.
pub fn is_known_constant(name: &str) -> bool {
    matches!(name,
        "pi" | "e" | "maxvalue" | "minvalue" |
        "positiveinfinity" | "negativeinfinity" | "nan" | "epsilon" |
        "empty" | "newline" | "true" | "false" |
        "completedtask"
    )
}

/// Return the .NET constructor table: bare type name → (host_module, host_constructor_func).
pub fn known_types() -> HashMap<String, (&'static str, &'static str)> {
    let mut m = HashMap::new();
    for (name, module, func) in &[
        // Collections
        ("list", "vybe:types", "listNew"),
        ("dictionary", "vybe:types", "dictNew"),
        ("queue", "vybe:types", "queueNew"),
        ("stack", "vybe:types", "stackNew"),
        ("hashset", "vybe:types", "hashSetNew"),
        ("arraylist", "vybe:types", "listNew"),
        ("hashtable", "vybe:types", "dictNew"),
        ("collection", "vybe:types", "listNew"),
        ("sortedlist", "vybe:types", "dictNew"),
        // Common types
        ("datetime", "vybe:types", "dateTimeNew"),
        ("stringbuilder", "vybe:types", "stringBuilderNew"),
        // Data
        ("datatable", "vybe:data", "dataTableNew"),
        ("dataset", "vybe:data", "dataSetNew"),
        // Drawing
        ("point", "vybe:drawing", "pointNew"),
        ("size", "vybe:drawing", "sizeNew"),
        ("sizef", "vybe:drawing", "sizeNew"),
        ("font", "vybe:drawing", "fontNew"),
        ("pen", "vybe:drawing", "penNew"),
        ("solidbrush", "vybe:drawing", "solidBrushNew"),
        ("color", "vybe:drawing", "colorFromName"),
        ("graphics", "vybe:drawing", "graphicsNew"),
        // Threading
        ("random", "vybe:threading", "randomNew"),
        ("stopwatch", "vybe:threading", "stopwatchNew"),
        // Database / Net
        ("sqlconnection", "vybe:database", "connect"),
        ("tcpclient", "vybe:net", "tcpConnect"),
        ("tcplistener", "vybe:net", "tcpListenerNew"),
        ("udpclient", "vybe:net", "udpNew"),
        ("streamreader", "vybe:net", "streamReaderNew"),
        ("streamwriter", "vybe:net", "streamWriterNew"),
        // Process
        ("processstartinfo", "vybe:types", "processStartInfoNew"),
        ("process", "vybe:types", "processNew"),
        // WinForms
        ("form", "vybe:gui", "newForm"),
    ] {
        m.insert(name.to_string(), (*module, *func));
    }
    m
}

// ─── Name shape helpers ──────────────────────────────────────────────────────
//
// The .NET surface — `System.Windows.Forms.Button`, `System.Windows.Forms.TextBox`,
// etc. — is one frontend on top of the canonical GUI vocabulary in
// `compiler_common::gui`. The .NET frontend's job is name resolution: take a
// .NET-shaped identifier and return the canonical control name. The actual
// emit (host fn naming, calling convention) lives in `gui.rs`.

/// Capitalise a lowercase WinForms control name to its proper casing.
/// Returns an empty string if the name is not a known control.
///
/// Thin wrapper over `compiler_common::gui::canonical_control_name`, kept for
/// backward compatibility with .NET-specific call sites. New code should call
/// into `gui.rs` directly. The .NET frontend assumes the source uses .NET
/// PascalCase, but the canonical name returned matches what other frontends
/// (MAUI, Flutter, Tkinter, …) would also produce.
pub fn capitalize_control_name(name: &str) -> String {
    crate::emitter::gui::canonical_control_name(name)
}

/// Data table / DataSet / DataAdapter — these are .NET BCL data types, NOT
/// GUI controls. They live in `known_types` because they're .NET-specific;
/// other framework frontends won't have them. Returns empty for non-data types.
pub fn capitalize_data_type(name: &str) -> String {
    match name {
        "dataset" => "DataSet",
        "datatable" => "DataTable",
        "dataadapter" => "DataAdapter",
        _ => return String::new(),
    }
    .to_string()
}
