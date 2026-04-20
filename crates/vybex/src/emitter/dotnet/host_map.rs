//! .NET → Vybe host mapping tables.
//!
//! The `.NET` BCL surface (`System.Console.WriteLine`, `Math.Sqrt`, etc.) is
//! exposed to user code via Vybe host functions (`wasi:cli::log`,
//! `vybe:math::sqrt`, etc.). This file owns BOTH translation tables that
//! make that work:
//!
//! 1. **`namespace_to_host_module`** — `system.console` → `wasi:cli`
//! 2. **`map_host_func`** — `(wasi:cli, writeline)` → `log`
//!
//! The two are kept together because the second table looks up the host
//! module from the first table when the resolver expands a dotted name into
//! a host call.
//!
//! GUI-specific mappings (`vybe:gui::new_Button` etc.) delegate to
//! `compiler_common::gui::canonical_control_name` so the canonical naming
//! lives in one place across all framework frontends.

/// Map a .NET namespace prefix (lowercased, dot-separated) to the Vybe host
/// module name. Returns the prefix itself if no explicit mapping exists.
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

/// Map a (host_module, dotnet_method_name) pair to the actual host function
/// name registered in the VM. Both inputs should already be lowercased.
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
        // The .NET surface uses Application.Run / Application.Exit, but the
        // canonical host fn names live in `compiler_common::gui`. Frontends
        // that aren't .NET-shaped (Tkinter `mainloop`, Flutter `runApp`, etc.)
        // will resolve to the SAME host fn names through their own frontend.
        ("vybe:gui", "application.run") => crate::emitter::gui::HOST_FN_RUN_APPLICATION.into(),
        ("vybe:gui", "run") => crate::emitter::gui::HOST_FN_RUN_APPLICATION.into(),
        ("vybe:gui", "exit") => crate::emitter::gui::HOST_FN_APP_EXIT.into(),
        ("vybe:gui", f) => {
            let canonical = crate::emitter::gui::canonical_control_name(f);
            if !canonical.is_empty() && canonical != f {
                crate::emitter::gui::host_fn_new_control(&canonical)
            } else {
                f.to_string()
            }
        }

        // ── Default: pass through ──
        (_, f) => f.to_string(),
    }
}
