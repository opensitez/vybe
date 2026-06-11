//! `System.Drawing` — value-type constructors and immediate-mode drawing.
//!
//! Two halves:
//!
//! - **`register(vm)`** — registered unconditionally. Installs the value
//!   constructors (`Pen`, `SolidBrush`, `Color`, `Point`, `Size`, `Font`,
//!   `Graphics`) plus no-op fallbacks for the `draw*`/`fill*` primitives.
//!   The fallbacks let non-GUI builds run drawing-using code without
//!   trapping on unresolved imports.
//!
//! - **`register_with_gui(vm, gui)`** — called from `gui::register` when
//!   the GUI backend is active. Overrides the `draw*`/`fill*` host fns
//!   with real implementations that capture each call as a `DrawCmd` on
//!   `GuiState::pending_draws`, keyed by the target control name. The
//!   form runner then drains the list each frame and replays the
//!   commands with `tiny_skia` (or any other rendering backend).
//!
//! All constructors live in `register` because they don't need
//! GuiState — they just produce dotnet `Object`s with the right shape.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::Object;
use vybe_bytecode::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    // Register Color as a global namespace with named colour constants.
    //
    // Each constant is an Object stamped with `r`, `g`, `b`, `a`
    // numeric fields (0-255). The dotnet `Graphics` body sequences
    // read these via `MethodOp::PushArgField(N, "r")` etc. when
    // building canvas calls — that's how `g.DrawLine(p, ...)` ends up
    // calling `vybe:gui::canvasSetStrokeColor(this, p.color.r,
    // p.color.g, p.color.b, p.color.a)` with real numeric args.
    {
        // (Name, R, G, B) — alpha is always 255 for the named constants.
        // Values match the .NET `KnownColor` enum where applicable.
        let palette: &[(&str, u8, u8, u8)] = &[
            ("Red", 220, 20, 60),
            ("Blue", 30, 144, 255),
            ("Green", 34, 139, 34),
            ("Black", 0, 0, 0),
            ("White", 255, 255, 255),
            ("Yellow", 255, 215, 0),
            ("Orange", 255, 140, 0),
            ("Purple", 128, 0, 128),
            ("Cyan", 0, 255, 255),
            ("Magenta", 255, 0, 255),
            ("Gray", 128, 128, 128),
            ("Brown", 139, 69, 19),
            ("Pink", 255, 192, 203),
            ("LightGray", 211, 211, 211),
            ("DarkGray", 169, 169, 169),
            ("Transparent", 0, 0, 0),
        ];

        let mut color_obj = Object::new();
        color_obj
            .properties
            .insert("__type".into(), Value::String(Arc::from("Color")));
        for (name, r, g, b) in palette {
            let alpha: u8 = if *name == "Transparent" { 0 } else { 255 };
            let mut c = Object::new();
            c.properties
                .insert("__type".into(), Value::String(Arc::from("Color")));
            c.properties
                .insert("name".into(), Value::String(Arc::from(*name)));
            c.properties.insert("r".into(), Value::F64(*r as f64));
            c.properties.insert("g".into(), Value::F64(*g as f64));
            c.properties.insert("b".into(), Value::F64(*b as f64));
            c.properties.insert("a".into(), Value::F64(alpha as f64));
            color_obj
                .properties
                .insert(name.to_lowercase(), Value::Object(Arc::new(Mutex::new(c))));
        }
        vm.globals.insert(
            "color".into(),
            Value::Object(Arc::new(Mutex::new(color_obj))),
        );
    }

    // New Point(x, y)
    vm.register_host_fn(
        "vybe:gui",
        "pointNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let x = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let y = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Point")));
            obj.properties.insert("x".into(), Value::F64(x));
            obj.properties.insert("y".into(), Value::F64(y));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // New Size(width, height)
    vm.register_host_fn(
        "vybe:gui",
        "sizeNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let w = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let h = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Size")));
            obj.properties.insert("width".into(), Value::F64(w));
            obj.properties.insert("height".into(), Value::F64(h));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // New Font(name, size)
    vm.register_host_fn(
        "vybe:gui",
        "fontNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let name = args
                .first()
                .map(|v| format!("{}", v))
                .unwrap_or_else(|| "Arial".into());
            let size = args.get(1).map(|v| v.as_f64()).unwrap_or(12.0);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Font")));
            obj.properties
                .insert("name".into(), Value::String(Arc::from(name.as_str())));
            obj.properties.insert("size".into(), Value::F64(size));
            obj.properties.insert("bold".into(), Value::Bool(false));
            obj.properties.insert("italic".into(), Value::Bool(false));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // New Pen(color, width)
    //
    // The pen state shape is `{ color, width, dashstyle, dashoffset }`.
    // Stroke methods on the framework wrapper layer (`Graphics.DrawLine`,
    // `Graphics.DrawRectangle`, …) read these fields via `PushArgField`
    // and forward them to canvas state mutations. `dashstyle` defaults
    // to 0 (Solid) and `dashoffset` to 0.0 — the framework wrapper's
    // setters mutate them in place when user code does
    // `pen.DashStyle = ...`.
    // `new Point(x, y)` — System.Drawing.Point value type. The GUI
    // property dispatch reads `.x` / `.y` from a `Value::Object` stored
    // at `location`, so we expose exactly those field names.
    vm.register_host_fn(
        "vybe:gui",
        "pointNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let x = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let y = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Point")));
            obj.properties.insert("x".into(), Value::F64(x));
            obj.properties.insert("y".into(), Value::F64(y));
            // Pascal-case aliases so code that reads `.X` / `.Y` (C#
            // idiomatic) resolves too; lowercase is what the GUI plumbing
            // actually reads.
            obj.properties.insert("X".into(), Value::F64(x));
            obj.properties.insert("Y".into(), Value::F64(y));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // `new Size(width, height)` — System.Drawing.Size. The GUI dispatch
    // reads `.width` / `.height` (lowercase) from the `size` property.
    vm.register_host_fn(
        "vybe:gui",
        "sizeNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let w = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let h = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Size")));
            obj.properties.insert("width".into(), Value::F64(w));
            obj.properties.insert("height".into(), Value::F64(h));
            obj.properties.insert("Width".into(), Value::F64(w));
            obj.properties.insert("Height".into(), Value::F64(h));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    vm.register_host_fn(
        "vybe:gui",
        "penNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let color = args.first().cloned().unwrap_or(Value::Null);
            let width = args.get(1).map(|v| v.as_f64()).unwrap_or(1.0);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Pen")));
            obj.properties.insert("color".into(), color);
            obj.properties.insert("width".into(), Value::F64(width));
            obj.properties.insert("dashstyle".into(), Value::F64(0.0));
            obj.properties.insert("dashoffset".into(), Value::F64(0.0));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // New SolidBrush(color)
    vm.register_host_fn(
        "vybe:gui",
        "solidBrushNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let color = args.first().cloned().unwrap_or(Value::Null);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("SolidBrush")));
            obj.properties.insert("color".into(), color);
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // Color.FromArgb(r, g, b) or Color.FromArgb(a, r, g, b)
    vm.register_host_fn(
        "vybe:gui",
        "color.fromargb",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let (a, r, g, b) = if args.len() >= 4 {
                (
                    args[0].as_f64() as u8,
                    args[1].as_f64() as u8,
                    args[2].as_f64() as u8,
                    args[3].as_f64() as u8,
                )
            } else if args.len() == 3 {
                (
                    255,
                    args[0].as_f64() as u8,
                    args[1].as_f64() as u8,
                    args[2].as_f64() as u8,
                )
            } else if args.len() == 1 {
                let val = args[0].as_f64() as u32;
                (
                    ((val >> 24) & 0xFF) as u8,
                    ((val >> 16) & 0xFF) as u8,
                    ((val >> 8) & 0xFF) as u8,
                    (val & 0xFF) as u8,
                )
            } else {
                (255, 0, 0, 0)
            };
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Color")));
            obj.properties.insert("r".into(), Value::F64(r as f64));
            obj.properties.insert("g".into(), Value::F64(g as f64));
            obj.properties.insert("b".into(), Value::F64(b as f64));
            obj.properties.insert("a".into(), Value::F64(a as f64));
            obj.properties.insert(
                "name".into(),
                Value::String(Arc::from(format!("#{:02X}{:02X}{:02X}", r, g, b).as_str())),
            );
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // ColorTranslator.FromHtml("#RRGGBB")
    vm.register_host_fn(
        "vybe:gui",
        "colortranslator.fromhtml",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let html = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let html = html.trim_start_matches('#');
            let (r, g, b) = if html.len() == 6 {
                (
                    u8::from_str_radix(&html[0..2], 16).unwrap_or(0),
                    u8::from_str_radix(&html[2..4], 16).unwrap_or(0),
                    u8::from_str_radix(&html[4..6], 16).unwrap_or(0),
                )
            } else {
                (0, 0, 0)
            };
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Color")));
            obj.properties.insert("r".into(), Value::F64(r as f64));
            obj.properties.insert("g".into(), Value::F64(g as f64));
            obj.properties.insert("b".into(), Value::F64(b as f64));
            obj.properties.insert("a".into(), Value::F64(255.0));
            obj.properties.insert(
                "name".into(),
                Value::String(Arc::from(format!("#{:02X}{:02X}{:02X}", r, g, b).as_str())),
            );
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // Color constants (Color.Red, Color.Blue, etc.) — stub as named objects
    vm.register_host_fn(
        "vybe:gui",
        "colorFromName",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let name = args
                .first()
                .map(|v| format!("{}", v))
                .unwrap_or_else(|| "Black".into());
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Color")));
            obj.properties
                .insert("name".into(), Value::String(Arc::from(name.as_str())));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // Graphics constructor. The dotnet `Graphics` class ctor calls this
    // host fn and copies `__control_name` from the result onto `this`. The
    // dotnet `Control.CreateGraphics` method (a `MethodTarget::DotnetCtor`)
    // routes through the dotnet ctor, which means by the time the user has
    // a `Graphics` instance in hand it carries the source control's name —
    // and the drawing host fns can read it back to route commands.
    //
    // The `__control_name` we stamp here is "" because this fn doesn't
    // know which control the dotnet `CreateGraphics` thunk was called
    // from; the canonical name is set by the user via the dotnet ctor's
    // identity copy step. The fallback "graphics" name is used by tests
    // that construct a `Graphics` directly.
    vm.register_host_fn(
        "vybe:gui",
        "graphicsNew",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Graphics")));
            obj.properties.insert(
                "__control_type".into(),
                Value::String(Arc::from("Graphics")),
            );
            obj.properties.insert(
                "__control_name".into(),
                Value::String(Arc::from("graphics")),
            );
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // NOTE: per-primitive `drawLine`/`fillRectangle`/etc. host fns
    // used to live here as no-op stubs. They were a placeholder for the
    // old "host knows how to draw" model. The new model routes
    // `Graphics.DrawLine(...)` through the `vybe:gui::canvas*` host
    // bridge (via dotnet `MethodTarget::Body` sequences), which calls
    // into the actual `vybe_widgets::canvas::Canvas` trait. This file
    // no longer registers any drawing primitives — only value-type
    // constructors (Pen, SolidBrush, Color, …).
    //
    // `Dispose` host fns also live in `vybe:gui::__ctrl_dispose` and
    // friends — see `modules/gui.rs`.

    // ── HatchBrush(BackgroundColor, ForegroundColor, HatchStyle) ──
    vm.register_host_fn(
        "vybe:gui",
        "hatchBrushNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let bg = args.first().cloned().unwrap_or(Value::Null);
            let fg = args.get(1).cloned().unwrap_or(Value::Null);
            let style = args.get(2).cloned().unwrap_or(Value::Null);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("HatchBrush")));
            obj.properties.insert("backgroundcolor".into(), bg);
            obj.properties.insert("foregroundcolor".into(), fg);
            obj.properties.insert("hatchstyle".into(), style);
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // ── LinearGradientBrush(Rect, Color1, Color2, Mode) ──
    vm.register_host_fn(
        "vybe:gui",
        "linearGradientBrushNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let rect = args.first().cloned().unwrap_or(Value::Null);
            let c1 = args.get(1).cloned().unwrap_or(Value::Null);
            let c2 = args.get(2).cloned().unwrap_or(Value::Null);
            let mode = args.get(3).cloned().unwrap_or(Value::Null);
            let mut obj = Object::new();
            obj.properties.insert(
                "__type".into(),
                Value::String(Arc::from("LinearGradientBrush")),
            );
            obj.properties.insert("rectangle".into(), rect);
            obj.properties.insert("startcolor".into(), c1);
            obj.properties.insert("endcolor".into(), c2);
            obj.properties.insert("wrapmode".into(), mode);
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );
}
