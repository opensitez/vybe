//! .NET namespace root recognition.
//!
//! "Is `Math` a variable or the start of a namespace/type chain?" The
//! resolver consults `is_namespace_root` to disambiguate. The set is
//! computed once at first access via `LazyLock`.
//!
//! Sibling files:
//! - `imports.rs` — the import LIST (what's auto-imported)
//! - `host_map.rs` — the .NET → host fn translation tables
//! - `resolver.rs` — uses `is_namespace_root` during dotted-name resolution

use std::collections::HashSet;
use std::sync::LazyLock;

/// The static set of namespace roots, computed once.
static NAMESPACE_ROOTS: LazyLock<HashSet<String>> = LazyLock::new(|| namespace_roots());

/// Check if a name is a known .NET namespace root.
pub fn is_namespace_root(name: &str) -> bool {
    NAMESPACE_ROOTS.contains(name)
}

/// Return the set of names that should be treated as namespace/type roots.
/// Public so language-specific extensions can build derived sets if needed.
pub fn namespace_roots() -> HashSet<String> {
    let mut s = HashSet::new();
    for name in &[
        // Top-level namespace roots
        "system", "microsoft", "vybe",
        // Direct .NET type names
        "math", "console", "convert", "strings", "array",
        "window", "file", "io", "directory", "application",
        "environment", "thread", "json", "color",
        "datetime", "stringbuilder", "process",
        "timespan", "guid", "point", "size", "font", "random",
        "path", "messagebox", "encoding",
        // WinForms enums
        "borderstyle", "formborderstyle", "contentalignment",
        "dialogresult", "messageboxbuttons", "messageboxicon",
        "keys", "dockstyle", "anchorstyles", "formstartposition",
        "formwindowstate",
        // System.Diagnostics types
        "stopwatch", "debug", "trace",
        // System.IO types
        "streamreader", "streamwriter", "filestream",
        "binaryreader", "binarywriter", "memorystream",
        // System.Net types
        "webrequest", "httpwebrequest", "webclient",
        "socket", "tcpclient", "tcplistener", "udpclient",
        // System.Threading types
        "task", "timer", "mutex", "semaphore",
        // System.Text types
        "regex", "match",
        // System.Collections types
        "list", "dictionary", "queue", "stack", "hashset",
        "arraylist", "hashtable", "sortedlist", "collection",
        // System.Data types
        "datatable", "dataset", "datarow", "datacolumn",
        "sqlconnection", "sqlcommand", "sqldatareader",
        "oledbconnection", "oledbcommand",
        // System.Xml types
        "xdocument", "xelement", "xmldocument",
        // System.Drawing types
        "pen", "solidbrush", "graphics", "bitmap", "image",
        "colortranslator", "systemcolors",
        // Primitive type names used as namespaces (Int32.Parse, etc.)
        "int", "integer", "long", "double", "single", "string", "boolean", "byte",
        "float", "bool", "object",
        "int32", "int64", "uint32",
    ] {
        s.insert(name.to_string());
    }
    s
}
