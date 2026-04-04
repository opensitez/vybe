//! System.Drawing — Point, Size, Font constructors

use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // Register Color as a global namespace with color constants
    {
        let mut color_obj = Object::new();
        color_obj.properties.insert("__type".into(), Value::String(Rc::from("Color")));
        for name in ["Red", "Blue", "Green", "Black", "White", "Yellow", "Orange",
                      "Purple", "Cyan", "Magenta", "Gray", "Brown", "Pink",
                      "LightGray", "DarkGray", "Transparent"] {
            let mut c = Object::new();
            c.properties.insert("__type".into(), Value::String(Rc::from("Color")));
            c.properties.insert("name".into(), Value::String(Rc::from(name)));
            color_obj.properties.insert(name.to_lowercase(), Value::Object(Rc::new(RefCell::new(c))));
        }
        vm.globals.insert("color".into(), Value::Object(Rc::new(RefCell::new(color_obj))));
    }

    // New Point(x, y)
    vm.register_host_fn("vybe:drawing", "pointNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let x = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
        let y = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Point")));
        obj.properties.insert("x".into(), Value::F64(x));
        obj.properties.insert("y".into(), Value::F64(y));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // New Size(width, height)
    vm.register_host_fn("vybe:drawing", "sizeNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let w = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
        let h = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Size")));
        obj.properties.insert("width".into(), Value::F64(w));
        obj.properties.insert("height".into(), Value::F64(h));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // New Font(name, size)
    vm.register_host_fn("vybe:drawing", "fontNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let name = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "Arial".into());
        let size = args.get(1).map(|v| v.as_f64()).unwrap_or(12.0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Font")));
        obj.properties.insert("name".into(), Value::String(Rc::from(name.as_str())));
        obj.properties.insert("size".into(), Value::F64(size));
        obj.properties.insert("bold".into(), Value::Bool(false));
        obj.properties.insert("italic".into(), Value::Bool(false));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // New Pen(color, width)
    vm.register_host_fn("vybe:drawing", "penNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let color = args.first().cloned().unwrap_or(Value::Null);
        let width = args.get(1).map(|v| v.as_f64()).unwrap_or(1.0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Pen")));
        obj.properties.insert("color".into(), color);
        obj.properties.insert("width".into(), Value::F64(width));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // New SolidBrush(color)
    vm.register_host_fn("vybe:drawing", "solidBrushNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let color = args.first().cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("SolidBrush")));
        obj.properties.insert("color".into(), color);
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Color.FromArgb(r, g, b) or Color.FromArgb(a, r, g, b)
    vm.register_host_fn("vybe:drawing", "color.fromargb", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let (a, r, g, b) = if args.len() >= 4 {
            (args[0].as_f64() as u8, args[1].as_f64() as u8, args[2].as_f64() as u8, args[3].as_f64() as u8)
        } else if args.len() == 3 {
            (255, args[0].as_f64() as u8, args[1].as_f64() as u8, args[2].as_f64() as u8)
        } else if args.len() == 1 {
            let val = args[0].as_f64() as u32;
            (((val >> 24) & 0xFF) as u8, ((val >> 16) & 0xFF) as u8, ((val >> 8) & 0xFF) as u8, (val & 0xFF) as u8)
        } else {
            (255, 0, 0, 0)
        };
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Color")));
        obj.properties.insert("r".into(), Value::F64(r as f64));
        obj.properties.insert("g".into(), Value::F64(g as f64));
        obj.properties.insert("b".into(), Value::F64(b as f64));
        obj.properties.insert("a".into(), Value::F64(a as f64));
        obj.properties.insert("name".into(), Value::String(Rc::from(format!("#{:02X}{:02X}{:02X}", r, g, b).as_str())));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // ColorTranslator.FromHtml("#RRGGBB")
    vm.register_host_fn("vybe:drawing", "colortranslator.fromhtml", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
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
        obj.properties.insert("__type".into(), Value::String(Rc::from("Color")));
        obj.properties.insert("r".into(), Value::F64(r as f64));
        obj.properties.insert("g".into(), Value::F64(g as f64));
        obj.properties.insert("b".into(), Value::F64(b as f64));
        obj.properties.insert("a".into(), Value::F64(255.0));
        obj.properties.insert("name".into(), Value::String(Rc::from(format!("#{:02X}{:02X}{:02X}", r, g, b).as_str())));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Color constants (Color.Red, Color.Blue, etc.) — stub as named objects
    vm.register_host_fn("vybe:drawing", "colorFromName", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let name = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "Black".into());
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Color")));
        obj.properties.insert("name".into(), Value::String(Rc::from(name.as_str())));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Graphics stub — CreateGraphics() returns a stub object
    vm.register_host_fn("vybe:drawing", "graphicsNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Graphics")));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // Drawing method stubs (DrawLine, FillRectangle, etc.) — no-ops for now
    for name in ["drawLine", "drawRectangle", "fillRectangle", "drawEllipse", "fillEllipse",
                 "drawString", "drawImage", "clear", "drawArc", "fillPolygon"] {
        vm.register_host_fn("vybe:drawing", name, Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::Null
        }));
    }
}
