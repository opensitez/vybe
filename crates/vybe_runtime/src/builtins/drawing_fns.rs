use crate::value::{RuntimeError, Value, ObjectData};
use std::collections::HashMap;
use std::rc::Rc;
use std::cell::RefCell;

/// Create a new System.Drawing.Color object
pub fn create_color_obj(name: &str, argb: u32) -> Value {
    let mut fields = HashMap::new();
    fields.insert("name".to_string(), Value::String(name.to_string()));
    fields.insert("r".to_string(), Value::Byte(((argb >> 16) & 0xFF) as u8));
    fields.insert("g".to_string(), Value::Byte(((argb >> 8) & 0xFF) as u8));
    fields.insert("b".to_string(), Value::Byte((argb & 0xFF) as u8));
    fields.insert("a".to_string(), Value::Byte(((argb >> 24) & 0xFF) as u8));
    
    let obj = ObjectData {
        class_name: "System.Drawing.Color".to_string(),
        fields,
        drawing_commands: Vec::new(),
    };
    Value::Object(Rc::new(RefCell::new(obj)))
}

/// System.Drawing.Color.FromArgb(r, g, b) or (a, r, g, b)
pub fn color_from_argb_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() == 3 {
        let r = args[0].as_integer()? as u8;
        let g = args[1].as_integer()? as u8;
        let b = args[2].as_integer()? as u8;
        let argb = 0xFF000000 | ((r as u32) << 16) | ((g as u32) << 8) | (b as u32);
        Ok(create_color_obj("", argb))
    } else if args.len() == 4 {
        let a = args[0].as_integer()? as u8;
        let r = args[1].as_integer()? as u8;
        let g = args[2].as_integer()? as u8;
        let b = args[3].as_integer()? as u8;
        let argb = ((a as u32) << 24) | ((r as u32) << 16) | ((g as u32) << 8) | (b as u32);
        Ok(create_color_obj("", argb))
    } else {
        Err(RuntimeError::Custom("Color.FromArgb requires 3 or 4 arguments".to_string()))
    }
}

/// System.Drawing.Pen(Color, [Width])
pub fn new_pen_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("New Pen requires at least color argument".to_string()));
    }
    
    let color = args[0].clone();
    let width = if args.len() > 1 {
        args[1].as_double()?
    } else {
        1.0
    };
    
    let mut fields = HashMap::new();
    fields.insert("color".to_string(), color);
    fields.insert("width".to_string(), Value::Double(width));
    
    let obj = ObjectData {
        class_name: "System.Drawing.Pen".to_string(),
        fields,
        drawing_commands: Vec::new(),
    };
    Ok(Value::Object(Rc::new(RefCell::new(obj))))
}

/// System.Drawing.SolidBrush(Color)
pub fn new_solid_brush_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("New SolidBrush requires color argument".to_string()));
    }
    
    let color = args[0].clone();
    
    let mut fields = HashMap::new();
    fields.insert("color".to_string(), color);
    
    let obj = ObjectData {
        class_name: "System.Drawing.SolidBrush".to_string(),
        fields,
        drawing_commands: Vec::new(),
    };
    Ok(Value::Object(Rc::new(RefCell::new(obj))))
}

/// System.Drawing.Graphics (Stub)
pub fn graphics_from_image_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    // Return a dummy Graphics object
    let obj = ObjectData {
        class_name: "System.Drawing.Graphics".to_string(),
        fields: HashMap::new(),
        drawing_commands: Vec::new(),
    };
    Ok(Value::Object(Rc::new(RefCell::new(obj))))
}
