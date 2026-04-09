//! Shared .NET BCL namespace resolution for all .NET compilers (VB, C#, etc.).
//!
//! The .NET Base Class Library uses the same namespace hierarchy regardless of
//! language: System.Threading.Thread.Sleep, System.Diagnostics.Stopwatch, etc.
//! This module provides a single source of truth so every compiler resolves
//! these identically.
//!
//! ## .NET Name Resolution Order
//!
//! When the compiler encounters a dotted name like `sw.ElapsedMilliseconds` or
//! `System.Threading.Thread.Sleep`, it must decide: is this a namespace chain
//! (static access) or an instance member access (local variable)?
//!
//! The resolution follows .NET semantics:
//! 1. **Locals first** — if the first part is a local variable, the rest is
//!    instance member access (struct_get chain on the object)
//! 2. **Module-level fields** — class/module fields (Me.field)
//! 3. **Fully-qualified namespace** — `System.Threading.Thread.Sleep`
//! 4. **Imports-resolved** — `Thread.Sleep` with `Imports System.Threading`
//! 5. **Known type static member** — `String.Format` as a type static method
//!
//! Compilers call `resolve_dotted_name()` which inspects the parts and returns
//! a `DottedResolution` telling the compiler exactly what bytecode to emit.

use std::collections::{HashMap, HashSet};

// ─── Resolution result ───────────────────────────────────────────────────────

/// The result of resolving a dotted name chain.
#[derive(Debug, Clone, PartialEq)]
pub enum DottedResolution {
    /// The first part is a local variable. The compiler should:
    /// - emit local_get for the variable
    /// - emit struct_get for each remaining part (instance member access)
    /// No call is implied; the caller decides whether to call or just access.
    InstanceMember {
        /// The local variable name (lowercased)
        local: String,
        /// The remaining parts after the local (lowercased member chain)
        members: Vec<String>,
    },

    /// Resolved to a host import via interface_imports (compile-time resolved).
    /// The compiler should emit call_import(module, func) directly.
    HostCall {
        module: String,
        func: String,
    },

    /// Resolved to a namespace object chain. The compiler should:
    /// - emit global_get for the root namespace
    /// - emit struct_get for each subsequent part
    /// The final value is a callable (for method calls) or a value (for property access).
    NamespaceAccess {
        /// The parts of the fully-qualified chain (lowercased)
        parts: Vec<String>,
    },

    /// A WinForms/layout no-op method. The compiler should emit null and skip.
    NoOp,

    /// Could not resolve — the compiler should fall back to its own logic.
    Unresolved,
}

/// Context provided by the compiler for name resolution.
/// This abstracts over VB/C# differences (different AST types, different scoping).
pub struct ResolutionContext<'a> {
    /// Check whether a name (lowercased) is a local variable in the current scope.
    pub is_local: &'a dyn Fn(&str) -> bool,
    /// Check whether a name (lowercased) is a field of the current class/module.
    pub is_class_field: &'a dyn Fn(&str) -> bool,
    /// Check whether a name is a user-defined class/module name.
    pub is_user_type: &'a dyn Fn(&str) -> bool,
    /// The active import list (e.g. ["system", "system.threading", ...])
    pub imports: &'a [String],
}

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

/// Resolve a dotted name chain following .NET resolution order.
///
/// `parts` is the member-access chain (e.g. ["sw", "ElapsedMilliseconds"] or
/// ["System", "Threading", "Thread", "Sleep"]). Already split by the caller;
/// NOT lowercased — this function handles casing.
///
/// `ctx` provides compiler-specific callbacks for scope lookup.
///
/// Returns a `DottedResolution` telling the compiler what to emit.
pub fn resolve_dotted_name(parts: &[&str], ctx: &ResolutionContext) -> DottedResolution {
    if parts.is_empty() {
        return DottedResolution::Unresolved;
    }

    let lower_parts: Vec<String> = parts.iter().map(|p| p.to_lowercase()).collect();
    let first = &lower_parts[0];

    // ── Step 0: No-op methods ────────────────────────────────────────────
    // If the LAST part is a known no-op method, short-circuit.
    if let Some(last) = lower_parts.last() {
        if is_noop_method(last) {
            return DottedResolution::NoOp;
        }
    }

    // ── Step 1: Local variable (highest priority) ────────────────────────
    // If the first part is a local variable, everything after it is instance
    // member access. This is how .NET works: locals shadow namespaces.
    if (ctx.is_local)(first) {
        return DottedResolution::InstanceMember {
            local: first.clone(),
            members: lower_parts[1..].to_vec(),
        };
    }

    // ── Step 2: Class field (Me.field implicit) ──────────────────────────
    if (ctx.is_class_field)(first) && lower_parts.len() > 1 {
        return DottedResolution::InstanceMember {
            local: first.clone(),
            members: lower_parts[1..].to_vec(),
        };
    }

    // ── Step 3: Fully-qualified namespace (System.X.Y.Z) ────────────────
    // Try to match the longest prefix against the import list.
    // This handles both `System.Threading.Thread.Sleep` (direct FQ) and
    // static type access through fully-qualified paths.
    if let Some(res) = try_resolve_via_imports(&lower_parts, ctx.imports) {
        return res;
    }

    // ── Step 4: Imports-resolved (bare type → prepend import prefix) ─────
    // e.g. "Thread.Sleep" with import "system.threading" →
    //      try resolving "system.threading.thread.sleep"
    if !is_namespace_root(first) {
        // Only try imports for names that aren't already namespace roots
        for import_path in ctx.imports {
            let mut expanded: Vec<String> = import_path.split('.').map(|s| s.to_string()).collect();
            expanded.extend(lower_parts.iter().cloned());
            let expanded_refs: Vec<&str> = expanded.iter().map(|s| s.as_str()).collect();
            if let Some(res) = try_resolve_via_imports_refs(&expanded_refs, ctx.imports) {
                return res;
            }
        }
    }

    // ── Step 5: Known namespace root but no import match ─────────────────
    // Fall back to namespace object chain (global_get → struct_get chain).
    // This handles enum values, static properties on namespace objects, etc.
    if is_namespace_root(first) {
        return DottedResolution::NamespaceAccess {
            parts: lower_parts,
        };
    }

    // ── Step 6: Try expanding with imports for namespace access ───────────
    // e.g. "Stopwatch.StartNew" with import "system.diagnostics"
    for import_path in ctx.imports {
        let mut expanded: Vec<String> = import_path.split('.').map(|s| s.to_string()).collect();
        expanded.extend(lower_parts.iter().cloned());
        let first_expanded = &expanded[0];
        if is_namespace_root(first_expanded) {
            return DottedResolution::NamespaceAccess {
                parts: expanded,
            };
        }
    }

    // ── Step 7: User-defined type static call ────────────────────────────
    if (ctx.is_user_type)(first) {
        return DottedResolution::NamespaceAccess {
            parts: lower_parts,
        };
    }

    DottedResolution::Unresolved
}

/// Check if a name is a known .NET namespace root.
pub fn is_namespace_root(name: &str) -> bool {
    NAMESPACE_ROOTS.contains(name)
}

/// Try to resolve a fully-qualified chain against the imports list.
/// Returns HostCall if it matches an import prefix, or NamespaceAccess
/// if the root is a known namespace.
fn try_resolve_via_imports(lower_parts: &[String], imports: &[String]) -> Option<DottedResolution> {
    let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
    try_resolve_via_imports_refs(&refs, imports)
}

fn try_resolve_via_imports_refs(lower_parts: &[&str], imports: &[String]) -> Option<DottedResolution> {
    if lower_parts.len() < 2 {
        return None;
    }

    // Try longest prefix first
    for prefix_len in (1..lower_parts.len()).rev() {
        let prefix = lower_parts[..prefix_len].join(".");
        if imports.contains(&prefix) {
            let func = lower_parts[prefix_len..].join(".");
            let module = namespace_to_host_module(&prefix);
            let mapped_func = map_host_func(module, &func);
            return Some(DottedResolution::HostCall {
                module: module.to_string(),
                func: mapped_func,
            });
        }
    }
    None
}

// The static set of namespace roots, computed once.
use std::sync::LazyLock;
static NAMESPACE_ROOTS: LazyLock<HashSet<String>> = LazyLock::new(|| namespace_roots());

// ─── Default interface imports ───────────────────────────────────────────────
// These are the .NET namespaces that are implicitly available (like having
// `Imports System` in VB or `using System;` in C#). Language compilers can
// extend this list with language-specific additions (e.g. "microsoft.visualbasic"
// for VB, "system.linq" for C#).

/// Return the default set of .NET namespace imports that every .NET compiler
/// should recognise.  Returned as a Vec so callers can `.extend()` with extras.
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

// ─── Namespace-to-host-module mapping ────────────────────────────────────────
// Maps a .NET namespace prefix to the Vybe host module that implements it.

/// Map a .NET namespace prefix (lowercased, dot-separated) to the Vybe host
/// module name.  Returns the prefix itself if no explicit mapping exists.
pub fn namespace_to_host_module<'a>(prefix: &'a str) -> &'a str {
    match prefix {
        "system.console" => "wasi:cli",
        "system.math" => "vybe:math",
        "system.convert" => "vybe:convert",
        "system.string" => "vybe:string",
        "system.array" => "vybe:array",
        "system.environment" => "wasi:cli",
        // IO
        "system.io" | "system.io.file" | "system.io.path" | "system.io.directory" => "wasi:filesystem",
        // Threading
        "system.threading.thread" => "wasi:clocks",
        "system.threading" | "system.threading.tasks" => "vybe:threading",
        // Diagnostics
        "system.diagnostics.process" => "vybe:types",
        "system.diagnostics.stopwatch" => "vybe:threading",
        "system.diagnostics.debug" | "system.diagnostics.trace" | "system.diagnostics" => "wasi:cli",
        // Net
        "system.net" => "wasi:http",
        "system.net.sockets" => "vybe:net",
        // Text
        "system.text.regularexpressions" => "vybe:regex",
        "system.text" => "vybe:string",
        // Collections
        "system.collections.generic" | "system.collections" => "vybe:types",
        // Data
        "system.data" | "system.data.sqlclient" | "system.data.oledb" => "vybe:data",
        // Security
        "system.security.cryptography" => "vybe:crypto",
        // XML
        "system.xml.linq" => "vybe:xml",
        // Drawing
        "system.drawing" => "vybe:drawing",
        // WinForms
        "system.windows.forms" => "vybe:gui",
        "application" => "vybe:gui",
        // VB-specific
        "microsoft.visualbasic" => "vybe:string",
        // Fallback
        _ => prefix,
    }
}

// ─── Method name mapping ─────────────────────────────────────────────────────
// Maps .NET method names (lowercased) to the actual host function names
// registered in the VM.

/// Map a (host_module, dotnet_method_name) pair to the actual host function
/// name.  Both inputs should already be lowercased.
pub fn map_host_func(module: &str, func: &str) -> String {
    match (module, func) {
        // ── Console ──
        ("wasi:cli", "writeline") => "log".into(),
        ("wasi:cli", "write") => "log".into(),
        ("wasi:cli", "readline") => "readLine".into(),
        ("wasi:cli", "error") => "error".into(),
        ("wasi:cli", "print") => "log".into(),
        ("wasi:cli", "assert") => "log".into(),

        // ── Math ──
        ("vybe:math", f) => f.to_string(),

        // ── Filesystem ──
        ("wasi:filesystem", "readalltext") => "readFile".into(),
        ("wasi:filesystem", "writealltext") => "writeFile".into(),
        ("wasi:filesystem", "appendalltext") => "appendFile".into(),
        ("wasi:filesystem", "exists") => "exists".into(),
        ("wasi:filesystem", "delete") => "remove".into(),
        ("wasi:filesystem", "copy") => "copy".into(),
        ("wasi:filesystem", "move") => "rename".into(),
        ("wasi:filesystem", "combine") => "pathCombine".into(),
        ("wasi:filesystem", "getfilename") => "pathGetFileName".into(),
        ("wasi:filesystem", "getextension") => "pathGetExtension".into(),
        ("wasi:filesystem", "getdirectoryname") => "pathGetDirectory".into(),
        ("wasi:filesystem", "getfilenamewithoutextension") => "pathGetFileNameWithoutExt".into(),
        ("wasi:filesystem", "changeextension") => "pathChangeExtension".into(),
        ("wasi:filesystem", "getfullpath") => "pathGetFullPath".into(),
        ("wasi:filesystem", "gettemppath") => "pathGetTempPath".into(),
        ("wasi:filesystem", "createdirectory") => "mkdir".into(),
        ("wasi:filesystem", "getfiles") => "listDir".into(),
        ("wasi:filesystem", "getcurrentdirectory") => "cwd".into(),

        // ── Convert ──
        ("vybe:convert", "toint32") => "cint".into(),
        ("vybe:convert", "todouble") => "cdbl".into(),
        ("vybe:convert", "tostring") => "toString".into(),
        ("vybe:convert", "toboolean") => "cbool".into(),
        ("vybe:convert", "todatetime") => "toString".into(),

        // ── Environment ──
        ("wasi:cli", "getenvironmentvariable") => "getEnv".into(),
        ("wasi:cli", "machinename") => "machineName".into(),
        ("wasi:cli", "currentdirectory") => "cwd".into(),

        // ── Threading ──
        ("wasi:clocks", "sleep") => "sleep".into(),

        // ── Diagnostics - Process ──
        ("vybe:types", "start") => "processStart".into(),
        ("vybe:types", "getcurrentprocess") => "processGetCurrent".into(),

        // ── Diagnostics - Stopwatch ──
        ("vybe:threading", "startnew") => "stopwatchNew".into(),

        // ── GUI / WinForms ──
        ("vybe:gui", "application.run") => "runApplication".into(),
        ("vybe:gui", "run") => "runApplication".into(),
        ("vybe:gui", "exit") => "appExit".into(),
        ("vybe:gui", f) => {
            let cap = capitalize_control_name(f);
            if !cap.is_empty() && cap != f {
                format!("new_{}", cap)
            } else {
                f.to_string()
            }
        }

        // ── Default: pass through ──
        (_, f) => f.to_string(),
    }
}

// ─── resolve_interface_call (legacy API, delegates to new resolver) ───────────
// Kept for backward compatibility. Callers should migrate to resolve_dotted_name().

/// Resolve a dotted .NET name to a `(module, function)` host import.
///
/// `parts` is the member-access chain split on `.`.
/// `interface_imports` is the active list of known namespace prefixes.
///
/// Returns `None` if no import prefix matches.
pub fn resolve_interface_call(parts: &[&str], interface_imports: &[String]) -> Option<(String, String)> {
    let lower_parts: Vec<String> = parts.iter().map(|p| p.to_lowercase()).collect();
    let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
    match try_resolve_via_imports_refs(&refs, interface_imports) {
        Some(DottedResolution::HostCall { module, func }) => Some((module, func)),
        _ => None,
    }
}

// ─── Known types (constructors) ──────────────────────────────────────────────
// Maps bare type name → (host_module, host_constructor_func) for `New <Type>()`

/// Return the .NET constructor table: bare type name → (module, func).
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

// ─── Namespace roots ─────────────────────────────────────────────────────────
// Names that the compiler should recognise as the start of a namespace or type
// chain rather than a variable reference.

/// Return the set of names that should be treated as namespace/type roots.
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

// ─── WinForms control name capitalisation ────────────────────────────────────

/// Capitalise a lowercase WinForms control name to its proper casing.
/// Returns an empty string if the name is not a known control.
pub fn capitalize_control_name(name: &str) -> String {
    match name {
        "button" => "Button",
        "label" => "Label",
        "textbox" => "TextBox",
        "checkbox" => "CheckBox",
        "radiobutton" => "RadioButton",
        "combobox" => "ComboBox",
        "listbox" => "ListBox",
        "panel" => "Panel",
        "groupbox" => "GroupBox",
        "tabcontrol" => "TabControl",
        "tabpage" => "TabPage",
        "datagridview" => "DataGridView",
        "progressbar" => "ProgressBar",
        "trackbar" => "TrackBar",
        "numericupdown" => "NumericUpDown",
        "datetimepicker" => "DateTimePicker",
        "richtextbox" => "RichTextBox",
        "picturebox" => "PictureBox",
        "menustrip" => "MenuStrip",
        "toolstrip" => "ToolStrip",
        "statusstrip" => "StatusStrip",
        "splitcontainer" => "SplitContainer",
        "flowlayoutpanel" => "FlowLayoutPanel",
        "tablelayoutpanel" => "TableLayoutPanel",
        "linklabel" => "LinkLabel",
        "maskedtextbox" => "MaskedTextBox",
        "listview" => "ListView",
        "webbrowser" => "WebBrowser",
        "monthcalendar" => "MonthCalendar",
        "contextmenustrip" => "ContextMenuStrip",
        "timer" => "Timer",
        "bindingsource" => "BindingSource",
        "tooltip" => "ToolTip",
        "imagelist" => "ImageList",
        "openfiledialog" => "OpenFileDialog",
        "savefiledialog" => "SaveFileDialog",
        "folderbrowserdialog" => "FolderBrowserDialog",
        "colordialog" => "ColorDialog",
        "fontdialog" => "FontDialog",
        "treeview" => "TreeView",
        "hscrollbar" => "HScrollBar",
        "vscrollbar" => "VScrollBar",
        "bindingnavigator" => "BindingNavigator",
        "dataset" => "DataSet",
        "datatable" => "DataTable",
        "dataadapter" => "DataAdapter",
        "notifyicon" => "NotifyIcon",
        "errorprovider" => "ErrorProvider",
        "helpprovider" => "HelpProvider",
        "backgroundworker" => "BackgroundWorker",
        _ => return String::new(),
    }
    .to_string()
}
