//! Vybe System Interface (VSI) modules.
//!
//! Each module registers host functions with (module, name) pairs on the VM,
//! following the WASI capability-based security model.
//!
//! Capabilities control which modules are available:
//! - Safe (always on): console, types, runtime, data, drawing
//! - Requires permission: fs, database, sockets, http, env, gui, crypto, xml
//! - Compiled-in (no host fns): threading — WASM threads proposal opcodes
//!   (thread.spawn / thread.join / atomics) emitted by compiler_common.

pub mod console;
// `vybe:array` retired — every former caller now routes through
// `ecma:array.*` (real ECMA-262 §23.1) or stdlib polyfills (`__vybe_*`
// chunks via `BuiltinEmit::Stdlib`). The module file is gone.
// `vybe:string` retired — VB Left/Mid/InStr/Format and PHP increment
// helpers now compile to `ecma:string.*` / direct opcodes / stdlib
// polyfills.
// `vybe:convert` retired — VB conversion semantics now compile to
// `ecma:number` / `ecma:string` / direct opcodes.
// `vybe:crypto` retired — md5/sha256 now flow through `web:crypto`
// (WHATWG SubtleCrypto) registered via `crate::web::register`.
pub mod fs;
pub mod clock;
pub mod env;
pub mod random;
// `vybe:http` (outbound HTTP client) moved to `crate::wasi::http` —
// it registers `wasi:http/{get,post,fetch}`, the spec-aligned namespace.
// `vybe:object` retired — `key in obj` flows through `ecma:object.hasOwn`
// (compiler normalises arg order); `a instanceof B` flows through
// `Op::REF_TEST` (WASM GC ref.test) for static type names + an inline
// __type/name string-compare fallback for dynamic RHS.
pub mod runtime;
pub mod database;
pub mod gui;
pub mod types;
pub mod sockets;
// `vybe:xml` retired — moved to `crate::web::dom_parser` exposing
// `web:dom-parser` (the WHATWG DOM Parsing & Serialization namespace).
pub mod data;
pub mod drawing;
pub mod canvas;
// `http_server` module retired — moved to `crate::node::http` (Node-aligned
// `node:http` namespace). The `register` fn is reachable via that path.
// Note: every `ecma:*` / `wasm:*` host module lives in the sibling
// `crate::ecma` and `crate::wasm` folders (not under `modules/`). See
// those `mod.rs` files for the full list. `register_with_capabilities`
// below calls `crate::ecma::register(vm)` and `crate::wasm::register(vm)`.

use vybe_bytecode::{VM, Value, HostContext};
use std::collections::HashSet;
use std::sync::{Arc, Mutex};

/// Capability flags for host module access.
/// Follows WASI's capability-based security model.
#[derive(Debug, Clone)]
pub struct Capabilities {
    granted: HashSet<Capability>,
}

/// Individual capabilities that can be granted or denied.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Capability {
    /// Console I/O (stdout, stderr). Safe for most contexts.
    Console,
    /// Filesystem read access.
    FileRead,
    /// Filesystem write access (implies FileRead).
    FileWrite,
    /// Network: outbound HTTP requests.
    Http,
    /// Network: TCP/UDP sockets (server + client).
    Sockets,
    /// Database connections (SQLite, MySQL, etc.).
    Database,
    /// Environment variables and process args.
    Environment,
    /// GUI / window creation.
    Gui,
    /// Threading / background tasks.
    Threading,
    /// Cryptographic operations.
    Crypto,
    /// System clock access (time, sleep).
    Clock,
    /// Random number generation.
    Random,
    /// XML parsing.
    Xml,
    /// HTTP server (binding ports, handling requests). Required for `vybex --serve`
    /// and any script calling `vybe:http/server.listen`.
    HttpServer,
    /// Spawning child processes (`node:child_process.{spawnSync, execSync,
    /// execFileSync}`, `node:process.kill`). Carries OS-level escape
    /// potential — gated separately from FileWrite.
    Process,
}
// `vybe:rt` retired — all dyn_* / get_prop / set_prop / new_object / array_* / typeof /
// global_get / struct_get_idx operations replaced by opcodes (Op::ADD, Op::STRUCT_GET,
// Op::TYPEOF, etc.). The module was a WASM bridge for the tree-walking interpreter;
// bytecode emits directly.

impl Capabilities {
    /// Full access — all capabilities granted. For trusted CLI usage.
    pub fn all() -> Self {
        use Capability::*;
        let mut granted = HashSet::new();
        for cap in [Console, FileRead, FileWrite, Http, Sockets, Database,
                    Environment, Gui, Threading, Crypto, Clock, Random, Xml, HttpServer, Process] {
            granted.insert(cap);
        }
        Capabilities { granted }
    }

    /// Safe subset — only pure computation, no I/O or side effects.
    /// Suitable for untrusted code (web playground, sandboxed eval).
    pub fn safe() -> Self {
        use Capability::*;
        let mut granted = HashSet::new();
        for cap in [Console, Clock, Random] {
            granted.insert(cap);
        }
        Capabilities { granted }
    }

    /// No capabilities — pure computation only.
    pub fn none() -> Self {
        Capabilities { granted: HashSet::new() }
    }

    /// Custom: start with none, add specific capabilities.
    pub fn with(caps: &[Capability]) -> Self {
        Capabilities { granted: caps.iter().copied().collect() }
    }

    pub fn has(&self, cap: Capability) -> bool {
        self.granted.contains(&cap)
    }

    pub fn grant(&mut self, cap: Capability) {
        self.granted.insert(cap);
    }

    pub fn revoke(&mut self, cap: Capability) {
        self.granted.remove(&cap);
    }
}

/// Register all standard VSI modules on a VM (no GUI).
/// All capabilities granted.
pub fn register_all(vm: &mut VM) {
    register_with_capabilities(vm, &Capabilities::all());

    // __debug_dump(obj) — print all properties of an object for debugging.
    // Useful in tests: `__debug_dump(myObj)` shows what's on the object.
    vm.register_host_fn("vybe:debug", "dump", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        use vybe_bytecode::value::ObjectKind;
        for (i, arg) in args.iter().enumerate() {
            match arg {
                Value::Object(obj) => {
                    let o = obj.lock().unwrap();
                    eprintln!("__debug_dump arg[{}]: Object (type_id={}) {{", i, o.type_id);
                    match &o.kind {
                        ObjectKind::Array(elems) => {
                            eprintln!("  [Array] length={}", elems.len());
                            for (j, elem) in elems.iter().enumerate().take(20) {
                                eprintln!("    [{}] = {}", j, elem);
                            }
                        }
                        ObjectKind::Function(f) => {
                            eprintln!("  [Function] name={:?}, arity={}, chunk={}",
                                f.name, f.arity, f.chunk_index);
                        }
                        ObjectKind::HostFunction(idx) => {
                            eprintln!("  [HostFunction] idx={}", idx);
                        }
                        ObjectKind::Map(m) => {
                            eprintln!("  [Map] size={}", m.len());
                            for (j, (k, v)) in m.iter().enumerate().take(20) {
                                eprintln!("    [{}] {} => {}", j, k, v);
                            }
                        }
                        ObjectKind::Set(s) => {
                            eprintln!("  [Set] size={}", s.len());
                            for (j, v) in s.iter().enumerate().take(20) {
                                eprintln!("    [{}] {}", j, v);
                            }
                        }
                        ObjectKind::ArrayBuffer(ab) => {
                            let bytes = ab.bytes.lock().unwrap();
                            eprintln!("  [ArrayBuffer] byteLength={} resizable={} detached={}",
                                bytes.len(), ab.resizable, ab.detached);
                        }
                        ObjectKind::TypedArray(ta) => {
                            eprintln!("  [TypedArray] elem={:?} length={} byteOffset={} byteLength={}",
                                ta.elem, ta.length, ta.byte_offset,
                                ta.length * ta.elem.bytes_per_element());
                        }
                        ObjectKind::Ordinary => {
                            eprintln!("  [Ordinary]");
                        }
                        ObjectKind::ModuleNamespace => {
                            eprintln!("  [ModuleNamespace] exports={}", o.properties.len());
                        }
                        ObjectKind::Continuation(cs) => {
                            let phase = *cs.state.lock().unwrap();
                            eprintln!("  [Continuation] phase={:?}", phase);
                        }
                    }
                    for (k, v) in &o.properties {
                        let v_str = format!("{}", v);
                        let v_short = if v_str.len() > 60 { format!("{}...", &v_str[..60]) } else { v_str };
                        eprintln!("  .{} = {}", k, v_short);
                    }
                    eprintln!("}}");
                }
                other => eprintln!("__debug_dump arg[{}]: {} ({})", i, other, other.type_tag()),
            }
        }
        Value::Null
    }));
    // Register no-op GUI stubs so compiled code that emits controlSetProperty/showForm/closeForm
    // doesn't fail with "Unresolved import" in non-GUI contexts.
    if vm.host_registry.get(&("vybe:gui".to_string(), "controlSetProperty".to_string())).is_none() {
        // Non-GUI stub that still mirrors the property write onto the object's
        // properties dict — this is essential because the dotnet class wrappers
        // emit setter chunks that call this fn, and user code (and tests) read
        // back the values via `obj.field`. The real GUI version of this fn
        // (`vybe_host::modules::gui::register::controlSetProperty`) ALSO writes
        // to a separate gui_state property store, which we skip here because
        // we have no `GuiState` to write into.
        vm.register_host_fn("vybe:gui", "controlSetProperty", Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let property = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                let val = args.get(2).cloned().unwrap_or(Value::Null);
                let prop_lower = property.to_lowercase();
                let mut o = obj.lock().unwrap();
                o.properties.insert(prop_lower.clone(), val.clone());
                if prop_lower == "name" {
                    o.properties.insert("__control_name".into(), val);
                }
            }
            Value::Null
        }));
        vm.register_host_fn("vybe:gui", "controlGetProperty", Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let property = args.get(1)
                    .map(|v| format!("{}", v).to_lowercase())
                    .unwrap_or_default();
                return obj.lock().unwrap().properties.get(&property).cloned().unwrap_or(Value::Null);
            }
            Value::Null
        }));
        vm.register_host_fn("vybe:gui", "setProperty", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "showForm", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "closeForm", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "showFormDialog", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "noop", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "runApplication", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "onEvent", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "controlsAdd", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "newControlsCollection", Box::new(|_ctx, args| {
            use vybe_bytecode::value::Object;
            let owner = args.first().cloned();
            let mut collection = Object::new_array(vec![]);
            collection.properties.insert("__type".into(), Value::String(Arc::from("ControlCollection")));
            if let Some(owner) = owner {
                collection.properties.insert("__owner".into(), owner);
            }
            collection.properties.insert("count".into(), Value::F64(0.0));
            Value::Object(Arc::new(Mutex::new(collection)))
        }));
        vm.register_host_fn("vybe:gui", "newComponentsCollection", Box::new(|_ctx, args| {
            use vybe_bytecode::value::Object;
            let owner = args.first().cloned();
            let mut collection = Object::new_array(vec![]);
            collection.properties.insert("__type".into(), Value::String(Arc::from("ComponentCollection")));
            if let Some(owner) = owner {
                collection.properties.insert("__owner".into(), owner);
            }
            collection.properties.insert("count".into(), Value::F64(0.0));
            Value::Object(Arc::new(Mutex::new(collection)))
        }));
        vm.register_host_fn("vybe:gui", "__collection_add", Box::new(|_ctx, args| {
            if let Some(Value::Object(collection)) = args.first() {
                let value = args.get(1).cloned().unwrap_or(Value::Null);
                let mut collection = collection.lock().unwrap();
                let mut len = None;
                if let vybe_bytecode::value::ObjectKind::Array(items) = &mut collection.kind {
                    if !items.iter().any(|existing| existing.eq(&value)) {
                        items.push(value);
                    }
                    len = Some(items.len());
                }
                if let Some(len) = len {
                    collection.properties.insert("count".into(), Value::F64(len as f64));
                    collection.properties.insert("length".into(), Value::F64(len as f64));
                }
            }
            Value::Null
        }));
        vm.register_host_fn("vybe:gui", "__collection_clear", Box::new(|_ctx, args| {
            if let Some(Value::Object(collection)) = args.first() {
                let mut collection = collection.lock().unwrap();
                if let vybe_bytecode::value::ObjectKind::Array(items) = &mut collection.kind {
                    items.clear();
                }
                collection.properties.insert("count".into(), Value::F64(0.0));
                collection.properties.insert("length".into(), Value::F64(0.0));
            }
            Value::Null
        }));
        vm.register_host_fn("vybe:gui", "__collection_contains", Box::new(|_ctx, args| {
            let Some(Value::Object(collection)) = args.first() else {
                return Value::Bool(false);
            };
            let needle = args.get(1).cloned().unwrap_or(Value::Null);
            let collection = collection.lock().unwrap();
            let contains = if let vybe_bytecode::value::ObjectKind::Array(items) = &collection.kind {
                items.iter().any(|existing| existing.eq(&needle))
            } else {
                false
            };
            Value::Bool(contains)
        }));
        vm.register_host_fn("vybe:gui", "newForm", Box::new(|_ctx, args| {
            use vybe_bytecode::value::Object;
            let title = args.first().map(|v| format!("{v}")).unwrap_or_default();
            let mut obj = Object::new();
            obj.properties.insert("__control_type".into(), Value::String(Arc::from("Form")));
            obj.properties.insert("text".into(), Value::String(Arc::from(title.as_str())));
            obj.properties.insert("name".into(), Value::String(Arc::from("form")));
            // Controls collection (no-op stub)
            let mut ctrls = Object::new_array(vec![]);
            ctrls.properties.insert("__type".into(), Value::String(Arc::from("ControlCollection")));
            ctrls.properties.insert("count".into(), Value::F64(0.0));
            obj.properties.insert("controls".into(), Value::Object(Arc::new(Mutex::new(ctrls))));
            let mut comps = Object::new_array(vec![]);
            comps.properties.insert("__type".into(), Value::String(Arc::from("ComponentCollection")));
            comps.properties.insert("count".into(), Value::F64(0.0));
            obj.properties.insert("components".into(), Value::Object(Arc::new(Mutex::new(comps))));
            Value::Object(Arc::new(Mutex::new(obj)))
        }));
        vm.register_host_fn("vybe:gui", "addHandler", Box::new(|_ctx, _| Value::Null));

        // ── Control / Form method stubs for the dotnet class wrappers ──
        // These are bound by `compiler_common::dotnet::classes::control::CONTROL_METHODS`
        // and `form::FORM_METHODS` as method thunks. Without the host
        // import target the VM would trap on unresolved import even
        // though no test actually exercises window lifecycle.
        for fn_name in &[
            "__ctrl_show", "__ctrl_hide", "__ctrl_focus", "__ctrl_close",
            "__ctrl_refresh", "__ctrl_invalidate", "__ctrl_update",
            "__ctrl_bring_to_front", "__ctrl_send_to_back", "__ctrl_dispose",
            "__form_activate", "__form_center_to_screen",
            "__dlg_showdialog", "__dlg_show",
        ] {
            vm.register_host_fn("vybe:gui", fn_name, Box::new(|_ctx, _| Value::Null));
        }

        // ── Canvas host fn stubs for the dotnet drawing wrappers ────────
        // The dotnet `Graphics` class compiles `DrawLine` etc. into
        // method thunks that call `vybe:gui::canvas*`. Tests run
        // through `register_all` (without `register_with_gui`), so the
        // real canvas impls in `modules::canvas` aren't installed —
        // these stubs let imports resolve and let drawing code run
        // without trapping. Tests that actually want to verify drawing
        // happened use `register_all_with_gui` which installs the real
        // impls that record into `GuiState.overlay_canvases`.
        //
        // The constructor `vybe:gui::getContext` returns a real canvas
        // context handle (a small Object stamped with __control_name)
        // so framework wrapper ctors can identity-copy off it.
        vm.register_host_fn("vybe:gui", "getContext", Box::new(|_ctx, args| {
            use vybe_bytecode::value::Object;
            let ctrl_name = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let mut o = Object::new();
            o.properties.insert("__type".into(), Value::String(Arc::from("CanvasContext")));
            o.properties.insert("__control_type".into(), Value::String(Arc::from("CanvasContext")));
            o.properties.insert("__control_name".into(), Value::String(Arc::from(ctrl_name.to_lowercase().as_str())));
            Value::Object(Arc::new(Mutex::new(o)))
        }));
        for fn_name in &[
            // Paint state
            "canvasSetFillColor", "canvasSetStrokeColor",
            "canvasSetLineWidth", "canvasSetMiterLimit", "canvasSetGlobalAlpha",
            "canvasSetLineCap", "canvasSetLineJoin", "canvasSetFont",
            // Path building
            "canvasBeginPath", "canvasClosePath",
            "canvasMoveTo", "canvasLineTo", "canvasQuadTo", "canvasBezierTo",
            "canvasArc", "canvasRect", "canvasEllipse",
            // Drawing
            "canvasFill", "canvasStroke",
            "canvasFillRect", "canvasStrokeRect", "canvasClearRect",
            "canvasFillText", "canvasStrokeText", "canvasDrawImage",
            // State stack
            "canvasSave", "canvasRestore",
            // Transforms
            "canvasTranslate", "canvasRotate", "canvasRotateDegrees", "canvasScale",
            "canvasTransform", "canvasResetTransform",
            // Convenience composites
            "canvasFillEllipseInRect", "canvasStrokeEllipseInRect", "canvasClearAll",
            "canvasStrokeArcInRect", "canvasFillPieInRect", "canvasStrokePieInRect",
            // Dashed strokes (fixed-arity setters)
            "canvasSetLineDashSolid", "canvasSetLineDash2",
            "canvasSetLineDash4", "canvasSetLineDash6",
            "canvasSetLineDashOffset", "canvasApplyPenDashStyle",
            // Clipping
            "canvasClip", "canvasResetClip",
        ] {
            vm.register_host_fn("vybe:gui", fn_name, Box::new(|_ctx, _| Value::Null));
        }

        // ── Per-control `new_<Type>` stubs for the dotnet class wrappers ──
        // The compiler_common::dotnet::classes layer emits ctor chunks that
        // call `vybe:gui::new_<ClassName>` for every concrete leaf. In test
        // / non-GUI contexts the real `gui::register` isn't called, so we
        // install no-op stubs that return a minimally-populated control
        // object — enough for the dotnet ctor's "transfer widget identity"
        // step to succeed without panicking.
        let dotnet_concrete_controls: &[&str] = &[
            "Form",
            // Buttons family
            "Button", "CheckBox", "RadioButton",
            // Text family
            "TextBox", "RichTextBox", "MaskedTextBox",
            // Labels
            "Label", "LinkLabel",
            // Lists
            "ComboBox", "ListBox", "ListView", "TreeView",
            // Containers
            "Panel", "GroupBox", "TabControl", "TabPage", "SplitContainer",
            "FlowLayoutPanel", "TableLayoutPanel",
            // Progress
            "ProgressBar", "TrackBar", "NumericUpDown",
            // Dates
            "DateTimePicker", "MonthCalendar",
            // Media
            "PictureBox", "WebBrowser",
            // Grids
            "DataGridView",
            // Strips
            "ToolStrip", "MenuStrip", "StatusStrip", "ContextMenuStrip",
            // Non-visual
            "Timer", "BindingSource", "ImageList", "ToolTip",
            "NotifyIcon", "ErrorProvider", "HelpProvider", "BackgroundWorker",
            // Dialogs
            "OpenFileDialog", "SaveFileDialog", "FontDialog", "ColorDialog",
            "FolderBrowserDialog",
            // Drawing
            "Canvas",
        ];
        for ct in dotnet_concrete_controls {
            let type_name = ct.to_string();
            vm.register_host_fn("vybe:gui", &format!("new_{}", ct), Box::new(move |_ctx, _args| {
                use vybe_bytecode::value::Object;
                use std::sync::atomic::{AtomicU32, Ordering};
                static COUNTER: AtomicU32 = AtomicU32::new(1);
                let id = COUNTER.fetch_add(1, Ordering::Relaxed);
                let name = format!("{}_{}", type_name.to_lowercase(), id);
                let mut obj = Object::new();
                obj.properties.insert("__control_type".into(), Value::String(Arc::from(type_name.as_str())));
                obj.properties.insert("__control_name".into(), Value::String(Arc::from(name.as_str())));
                obj.properties.insert("__type".into(), Value::String(Arc::from(type_name.as_str())));
                obj.properties.insert("name".into(), Value::String(Arc::from(name.as_str())));
                Value::Object(Arc::new(Mutex::new(obj)))
            }));
        }
    }
    // DO NOT call setup_namespaces here — tests override host fns after register_all.
    // setup_namespaces must be called AFTER all host fn registrations.
}

/// Register modules based on granted capabilities.
pub fn register_with_capabilities(vm: &mut VM, caps: &Capabilities) {
    // Always registered — pure computation, no security risk.
    // Note: `ecma:math` moved to `crate::ecma::math` (registered below
    // via `crate::ecma::register`). Legacy `vybe:string` / `vybe:json`
    // / `vybe:object` / etc. still live under `modules/` until their
    // callers migrate to `ecma:*`.
    // `vybe:json` retired — JSON.parse / JSON.stringify both flow through
    // `ecma:json` (registered via `crate::ecma::register` below).
        // RegExp + String.prototype regex methods flow through `ecma:regexp`
        // (registered via `crate::ecma::register` below). Pattern-first
        // language conventions (PHP preg_*, Python re.*, VB Regex.*, .NET
        // System.Text.RegularExpressions.Regex) are bridged via stdlib adapter
        // chunks (`__stdlib_regex_*_pat_first`).
    // `vybe:collections` retired — JS Map/Set/WeakMap/WeakSet now flow
    // through `ecma:map` / `ecma:set` / `ecma:weakmap` / `ecma:weakset`
    // (registered via `crate::ecma::register` below). TypeRegistry
    // dispatch in `builtin_types.rs` points at the same fns.
    runtime::register(vm);
    types::register(vm);
    data::register(vm);
    drawing::register(vm);

    // ecma:* — JS runtime (ECMA-262 mirror). All pure computation
    // except `ecma:date`, which is gated under Clock below.
    crate::ecma::register(vm);
    // web:* — WHATWG / W3C web platform APIs (crypto, url, encoding,
    // fetch, timers). Some entries hit the network/disk and ought to
    // be capability-gated; today it's all-on for parity with how
    // browsers expose them.
    crate::web::register(vm);
    // wasm:* — real WebAssembly CG proposal host imports
    // (js-string-builtins + stage-1 js-primitive-builtins).
    crate::wasm::register(vm);

    // Capability-gated modules
    if caps.has(Capability::Console) {
        console::register(vm);
    }
    if caps.has(Capability::Clock) {
        clock::register(vm);
        // ecma:date reads the system clock — gated under Clock.
        crate::ecma::date::register(vm);
    }
    if caps.has(Capability::Random) {
        random::register(vm);
    }
    if caps.has(Capability::FileRead) || caps.has(Capability::FileWrite) {
        fs::register(vm);
        // node:fs — Node.js built-in filesystem surface, gated under
        // the same capability since reads/writes the same disk.
        crate::node::fs::register(vm);
        // wasi:filesystem 0.2.8 — real descriptor-based proposal.
        // Lives alongside the legacy Vybe-shim filesystem until the
        // shim is renamed to `vybe:fs`.
        crate::wasi::filesystem::register(vm);
    }
    // node:os, node:path, node:process — read-only system info / pure
    // computation; available regardless of capability.
    crate::node::register_always_on(vm);
    if caps.has(Capability::Process) {
        // node:child_process — gated separately since spawning OS
        // commands is a stronger capability than fs read/write.
        crate::node::child_process::register(vm);
    }
    // `dotnet:io` host fns retired — `StreamReader` / `StreamWriter` lower
    // at compile time through `emitter::dotnet::core::stream_io_adapter`
    // composing `node:fs.{read,write}FileSync`. No host module to register.
    if caps.has(Capability::Environment) {
        env::register(vm);
    }
    if caps.has(Capability::Http) {
        crate::wasi::http::register(vm);
    }
    if caps.has(Capability::Sockets) {
        sockets::register(vm);
        // `register_dotnet_net` / `register_dotnet_sockets` retired —
        // .NET Dns / TcpClient / TcpListener / UdpClient now lower to
        // `wasi:sockets/*` via `emitter::dotnet::core::sockets_adapter`.
    }
    if caps.has(Capability::Database) {
        database::register(vm);
    }
    // Capability::Threading — no host-fn module to register. Thread spawn /
    // join / atomics compile to WASM opcodes (Op::THREAD_SPAWN, etc.), and
    // Thread.Sleep uses wasi:clocks/sleep. The capability flag remains as a
    // gate for future wasi-threads primitives.
    // `vybe:crypto` retired — md5/sha256 flow through `web:crypto`
    // (registered unconditionally via `crate::web::register` above).
    // The Crypto capability gate remains for future wasi:crypto/* primitives.
    let _ = caps.has(Capability::Crypto);
    // `Xml` capability now gates access to `web:dom-parser` (registered
    // unconditionally via `crate::web::register` above). The flag stays
    // for future wasi:xml-style proposals or capability-restricted
    // sandboxes; today it's a no-op since DOM parsing is pure
    // computation.
    let _ = caps.has(Capability::Xml);
    if caps.has(Capability::HttpServer) {
        crate::node::http::register(vm);
    }

    // Set up namespace objects, type registry
    crate::namespaces::setup_namespaces(vm);
    crate::builtin_types::register_all(vm);

    // Polyfill: override stdlib `__vybe_*` globals with the just-registered
    // native host fns. After this, code that emits `global_get __vybe_pow +
    // call_ref` calls the native pow instead of the bundled stdlib bytecode
    // fallback. This is the runtime half of the polyfill pattern: stdlib
    // provides a portable WASM implementation, Vybe replaces it with a fast
    // native one. Non-Vybe runtimes simply keep the stdlib version.
    override_stdlib_globals_with_host_fns(vm);
}

/// Walk `IMPORT_ALIASES` and overwrite each `__vybe_*` global with the
/// corresponding registered host fn (if any). Idempotent — safe to call
/// multiple times.
pub fn override_stdlib_globals_with_host_fns(vm: &mut VM) {
    for &(module, name, global_name) in crate::stdlib_aliases::IMPORT_ALIASES {
        if let Some(&idx) = vm.host_registry.get(&(module.to_string(), name.to_string())) {
            if let Some(host_val) = vm.func_table.get(idx).cloned() {
                vm.globals.insert(global_name.to_string(), host_val);
            }
        }
    }
}

/// Register all standard VSI modules + GUI module.
/// Returns the shared GuiState — pass it to the form launcher.
#[cfg(feature = "gui")]
pub fn register_all_with_gui(
    vm: &mut VM,
) -> std::sync::Arc<std::sync::Mutex<crate::gui_state::GuiState>> {
    let gui = std::sync::Arc::new(std::sync::Mutex::new(crate::gui_state::GuiState::new()));
    register_all(vm);
    gui::register(vm, gui.clone());
    canvas::register(vm, gui.clone());
    // DO NOT call setup_namespaces here — callers do it after all overrides.
    gui
}

/// Register with capabilities + GUI.
/// Returns the shared GuiState.
#[cfg(feature = "gui")]
pub fn register_with_capabilities_and_gui(
    vm: &mut VM,
    caps: &Capabilities,
) -> std::sync::Arc<std::sync::Mutex<crate::gui_state::GuiState>> {
    let gui = std::sync::Arc::new(std::sync::Mutex::new(crate::gui_state::GuiState::new()));
    register_with_capabilities(vm, caps);
    if caps.has(Capability::Gui) {
        gui::register(vm, gui.clone());
        canvas::register(vm, gui.clone());
        crate::namespaces::setup_namespaces(vm);
    }
    gui
}
