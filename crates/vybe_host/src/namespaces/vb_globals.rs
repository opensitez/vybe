//! VB6-era globals: `App` and `Screen`.
//!
//! Legacy VB6 / VBA exposes `App.Path`, `App.Title`, `App.EXEName`, and
//! `Screen.Width`, `Screen.Height` as top-level object-like globals with
//! pre-computed properties. Modern VB.NET replaces these with
//! `My.Application` and `My.Computer.Screen`, but VB6 source still
//! accesses them via the short names.
//!
//! We register them as plain namespace objects with the values snapshot at
//! VM setup time. That's enough for every test + example that references
//! them, and keeps user code ignorant of the host.
//!
//! For values that require a real display (`Screen.Width`, `Screen.Height`)
//! we fall back to 1920×1080 when no display backend is active. A GUI
//! backend that knows the real resolution can overwrite the values before
//! running user code.

use super::*;
use std::path::PathBuf;

pub fn register(vm: &mut VM) {
    // ── App ─────────────────────────────────────────────────────────────
    //
    // `Path`      — directory of the running exe (no trailing slash)
    // `EXEName`   — exe file name without the `.exe` extension
    // `Title`     — same as EXEName unless overridden
    // `HInstance` — in VB6 this was the module handle; here we use 0
    // `PrevInstance` — always False (no previous instance in a modern OS)
    // `NonModalAllowed` — VB6 legacy flag, always True
    let app = ensure_namespace(vm, &["App"]);

    let (exe_dir, exe_name) = exe_path_and_name();
    set_prop(&app, "path",           Value::String(Arc::from(exe_dir.as_str())));
    set_prop(&app, "exename",        Value::String(Arc::from(exe_name.as_str())));
    set_prop(&app, "title",          Value::String(Arc::from(exe_name.as_str())));
    set_prop(&app, "hinstance",      Value::F64(0.0));
    set_prop(&app, "previnstance",   Value::F64(0.0));  // False
    set_prop(&app, "nonmodalallowed", Value::F64(1.0)); // True

    // ── Screen ──────────────────────────────────────────────────────────
    //
    // VB6 exposed `Screen.Width` / `Screen.Height` in twips (1/1440 in).
    // Real WinForms `Screen.PrimaryScreen.WorkingArea.Width` is in pixels.
    // We return pixel values because that's what user code actually wants,
    // and because the twips scaling is a compatibility shim nobody expects
    // in new code.
    //
    // The defaults (1920×1080) get overwritten by the GUI backend at
    // startup via `set_property` if a real display is available.
    let screen = ensure_namespace(vm, &["Screen"]);
    set_prop(&screen, "width",  Value::F64(1920.0));
    set_prop(&screen, "height", Value::F64(1080.0));
    set_prop(&screen, "twipsperpixelx", Value::F64(15.0));
    set_prop(&screen, "twipsperpixely", Value::F64(15.0));
}

/// Return `(directory, stem)` of the running executable.
///
/// If `std::env::current_exe` fails (rare — happens on exotic platforms or
/// when the binary was deleted out from under us), falls back to the
/// current working directory and an empty stem rather than panicking.
fn exe_path_and_name() -> (String, String) {
    let exe: PathBuf = std::env::current_exe().unwrap_or_else(|_| {
        std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."))
    });
    let dir = exe.parent()
        .map(|p| p.to_string_lossy().into_owned())
        .unwrap_or_default();
    let stem = exe.file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_default();
    (dir, stem)
}
