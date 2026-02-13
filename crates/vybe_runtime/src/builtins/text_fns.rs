use crate::value::{RuntimeError, Value};
use std::cell::RefCell;
use std::rc::Rc;

// ═══════════════════════════════════════════════════════════════════════════
// StringBuilder Implementation
// ═══════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub struct StringBuilder {
    buffer: String,
}

impl StringBuilder {
    pub fn new() -> Self {
        StringBuilder {
            buffer: String::new(),
        }
    }

    pub fn with_capacity(capacity: usize) -> Self {
        StringBuilder {
            buffer: String::with_capacity(capacity),
        }
    }

    pub fn append(&mut self, value: &str) {
        self.buffer.push_str(value);
    }

    pub fn append_line(&mut self, value: &str) {
        self.buffer.push_str(value);
        self.buffer.push('\n');
    }

    pub fn insert(&mut self, index: usize, value: &str) {
        if index <= self.buffer.len() {
            self.buffer.insert_str(index, value);
        }
    }

    pub fn remove(&mut self, start: usize, length: usize) {
        if start < self.buffer.len() {
            let end = (start + length).min(self.buffer.len());
            self.buffer.replace_range(start..end, "");
        }
    }

    pub fn replace(&mut self, old: &str, new: &str) {
        self.buffer = self.buffer.replace(old, new);
    }

    pub fn clear(&mut self) {
        self.buffer.clear();
    }

    pub fn to_string(&self) -> String {
        self.buffer.clone()
    }

    pub fn length(&self) -> usize {
        self.buffer.len()
    }

    pub fn capacity(&self) -> usize {
        self.buffer.capacity()
    }
}

/// StringBuilder.New() - Create new StringBuilder
pub fn stringbuilder_new_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    let sb = if args.is_empty() {
        StringBuilder::new()
    } else if let Ok(cap) = args[0].as_integer() {
        // New StringBuilder(capacity) — integer argument = capacity
        StringBuilder::with_capacity(cap as usize)
    } else {
        // New StringBuilder(initialValue) — string argument = initial content
        let initial = args[0].as_string();
        let mut sb = StringBuilder::new();
        sb.append(&initial);
        sb
    };
    
    use crate::value::ObjectData;
    let mut fields = std::collections::HashMap::new();
    fields.insert("__type".to_string(), Value::String("StringBuilder".to_string()));
    fields.insert("__data".to_string(), Value::String(sb.buffer.clone()));
    fields.insert("length".to_string(), Value::Integer(sb.buffer.len() as i32));
    fields.insert("capacity".to_string(), Value::Integer(sb.buffer.capacity() as i32));
    fields.insert("maxcapacity".to_string(), Value::Integer(i32::MAX));
    
    Ok(Value::Object(Rc::new(RefCell::new(ObjectData {
        class_name: "StringBuilder".to_string(),
        fields,
    }))))
}

/// StringBuilder methods dispatcher
pub fn stringbuilder_method_fn(method: &str, obj: &Value, args: &[Value]) -> Result<Value, RuntimeError> {
    if let Value::Object(obj_rc) = obj {
        let mut obj_data = obj_rc.borrow_mut();
        
        // Get current buffer
        let buffer = obj_data.fields.get("__data")
            .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
            .unwrap_or_default();
        
        let mut sb = StringBuilder { buffer };
        
        // Helper: sync length/capacity properties and return Self for chaining
        let return_self = |obj_data: &mut std::cell::RefMut<crate::value::ObjectData>, buffer: String| -> Result<Value, RuntimeError> {
            let len = buffer.len() as i32;
            let cap = buffer.capacity() as i32;
            obj_data.fields.insert("__data".to_string(), Value::String(buffer));
            obj_data.fields.insert("length".to_string(), Value::Integer(len));
            obj_data.fields.insert("capacity".to_string(), Value::Integer(cap));
            Ok(obj.clone())
        };

        match method.to_lowercase().as_str() {
            "append" => {
                if !args.is_empty() {
                    sb.append(&args[0].as_string());
                }
                return_self(&mut obj_data, sb.buffer)
            }
            "appendline" => {
                if args.is_empty() {
                    sb.buffer.push('\n');
                } else {
                    sb.append_line(&args[0].as_string());
                }
                return_self(&mut obj_data, sb.buffer)
            }
            "appendformat" => {
                // AppendFormat(format, args...) — simplified: just concatenate
                if !args.is_empty() {
                    let fmt = args[0].as_string();
                    // Simple {0}, {1}, ... substitution
                    let mut result = fmt.clone();
                    for (i, arg) in args[1..].iter().enumerate() {
                        result = result.replace(&format!("{{{}}}", i), &arg.as_string());
                    }
                    sb.append(&result);
                }
                return_self(&mut obj_data, sb.buffer)
            }
            "insert" => {
                if args.len() >= 2 {
                    let index = args[0].as_integer()? as usize;
                    let value = args[1].as_string();
                    sb.insert(index, &value);
                }
                return_self(&mut obj_data, sb.buffer)
            }
            "remove" => {
                if args.len() >= 2 {
                    let start = args[0].as_integer()? as usize;
                    let length = args[1].as_integer()? as usize;
                    sb.remove(start, length);
                }
                return_self(&mut obj_data, sb.buffer)
            }
            "replace" => {
                if args.len() >= 2 {
                    let old = args[0].as_string();
                    let new_val = args[1].as_string();
                    sb.replace(&old, &new_val);
                }
                return_self(&mut obj_data, sb.buffer)
            }
            "clear" => {
                sb.clear();
                return_self(&mut obj_data, sb.buffer)
            }
            "tostring" => {
                Ok(Value::String(sb.to_string()))
            }
            "length" => {
                Ok(Value::Integer(sb.length() as i32))
            }
            "capacity" => {
                Ok(Value::Integer(sb.capacity() as i32))
            }
            "ensurecapacity" => {
                let cap = args.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                sb.buffer.reserve(cap.saturating_sub(sb.buffer.capacity()));
                return_self(&mut obj_data, sb.buffer)
            }
            "chars" => {
                let idx = args.get(0).and_then(|v| v.as_integer().ok()).unwrap_or(0) as usize;
                let c = sb.buffer.chars().nth(idx).unwrap_or('\0');
                Ok(Value::Char(c))
            }
            "copyto" => {
                // CopyTo(sourceIndex, destination As Char(), destinationIndex, count)
                if args.len() >= 4 {
                    let source_index = args[0].as_integer()? as usize;
                    let dest_index = args[2].as_integer()? as usize;
                    let count = args[3].as_integer()? as usize;
                    let chars: Vec<char> = sb.buffer.chars().collect();
                    if let Value::Array(ref dest_arr) = args[1] {
                        let mut new_arr = dest_arr.clone();
                        for i in 0..count {
                            if source_index + i < chars.len() && dest_index + i < new_arr.len() {
                                new_arr[dest_index + i] = Value::Char(chars[source_index + i]);
                            }
                        }
                        return Ok(Value::Array(new_arr));
                    }
                }
                Ok(Value::Nothing)
            }
            "equals" => {
                if let Some(Value::Object(other_rc)) = args.get(0) {
                    let other_buf = other_rc.borrow().fields.get("__data")
                        .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                        .unwrap_or_default();
                    Ok(Value::Boolean(sb.buffer == other_buf))
                } else {
                    Ok(Value::Boolean(false))
                }
            }
            _ => Err(RuntimeError::Custom(format!("Unknown StringBuilder method: {}", method)))
        }
    } else {
        Err(RuntimeError::Custom("Not a StringBuilder object".to_string()))
    }
}

// ═══════════════════════════════════════════════════════════════════════════
// Encoding Implementation
// ═══════════════════════════════════════════════════════════════════════════

/// Encoding.ASCII.GetBytes(string) - Convert string to ASCII bytes
pub fn encoding_ascii_getbytes_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("GetBytes requires a string argument".to_string()));
    }
    let text = args[0].as_string();
    let bytes: Vec<Value> = text.as_bytes().iter().map(|&b| Value::Integer(b as i32)).collect();
    Ok(Value::Array(bytes))
}

/// Encoding.ASCII.GetString(bytes) - Convert ASCII bytes to string
pub fn encoding_ascii_getstring_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("GetString requires a byte array argument".to_string()));
    }
    
    if let Value::Array(arr) = &args[0] {
        let bytes: Result<Vec<u8>, _> = arr.iter().map(|v| {
            v.as_integer().map(|i| i as u8)
        }).collect();
        
        match bytes {
            Ok(b) => Ok(Value::String(String::from_utf8_lossy(&b).to_string())),
            Err(e) => Err(e),
        }
    } else {
        Err(RuntimeError::Custom("GetString requires a byte array".to_string()))
    }
}

/// Encoding.UTF8.GetBytes(string) - Convert string to UTF-8 bytes
pub fn encoding_utf8_getbytes_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("GetBytes requires a string argument".to_string()));
    }
    let text = args[0].as_string();
    let bytes: Vec<Value> = text.as_bytes().iter().map(|&b| Value::Integer(b as i32)).collect();
    Ok(Value::Array(bytes))
}

/// Encoding.UTF8.GetString(bytes) - Convert UTF-8 bytes to string
pub fn encoding_utf8_getstring_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("GetString requires a byte array argument".to_string()));
    }
    
    if let Value::Array(arr) = &args[0] {
        let bytes: Result<Vec<u8>, _> = arr.iter().map(|v| {
            v.as_integer().map(|i| i as u8)
        }).collect();
        
        match bytes {
            Ok(b) => match String::from_utf8(b.clone()) {
                Ok(s) => Ok(Value::String(s)),
                Err(_) => Ok(Value::String(String::from_utf8_lossy(&b).to_string())),
            },
            Err(e) => Err(e),
        }
    } else {
        Err(RuntimeError::Custom("GetString requires a byte array".to_string()))
    }
}

/// Encoding.Unicode.GetBytes(string) - Convert string to UTF-16 bytes
pub fn encoding_unicode_getbytes_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("GetBytes requires a string argument".to_string()));
    }
    let text = args[0].as_string();
    let bytes: Vec<Value> = text.encode_utf16()
        .flat_map(|u| vec![(u & 0xFF) as u8, ((u >> 8) & 0xFF) as u8])
        .map(|b| Value::Integer(b as i32))
        .collect();
    Ok(Value::Array(bytes))
}

/// Encoding.Unicode.GetString(bytes) - Convert UTF-16 bytes to string
pub fn encoding_unicode_getstring_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("GetString requires a byte array argument".to_string()));
    }
    
    if let Value::Array(arr) = &args[0] {
        let bytes: Result<Vec<u8>, _> = arr.iter().map(|v| {
            v.as_integer().map(|i| i as u8)
        }).collect();
        
        match bytes {
            Ok(b) => {
                let u16_vec: Vec<u16> = b.chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                    .collect();
                Ok(Value::String(String::from_utf16_lossy(&u16_vec)))
            },
            Err(e) => Err(e),
        }
    } else {
        Err(RuntimeError::Custom("GetString requires a byte array".to_string()))
    }
}

/// Encoding.Default.GetBytes(string) - Convert string to system default encoding (UTF-8)
pub fn encoding_default_getbytes_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    // System default is UTF-8 on modern systems
    encoding_utf8_getbytes_fn(args)
}

/// Encoding.Default.GetString(bytes) - Convert system default encoding to string (UTF-8)
pub fn encoding_default_getstring_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    // System default is UTF-8 on modern systems
    encoding_utf8_getstring_fn(args)
}

/// Encoding.GetEncoding(name) - Get encoding by name
pub fn encoding_getencoding_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("GetEncoding requires an encoding name".to_string()));
    }
    
    let encoding_name = args[0].as_string().to_lowercase();
    
    use crate::value::ObjectData;
    let mut fields = std::collections::HashMap::new();
    
    match encoding_name.as_str() {
        "utf-8" | "utf8" => {
            fields.insert("name".to_string(), Value::String("UTF-8".to_string()));
            fields.insert("codepage".to_string(), Value::Integer(65001));
        }
        "ascii" | "us-ascii" => {
            fields.insert("name".to_string(), Value::String("ASCII".to_string()));
            fields.insert("codepage".to_string(), Value::Integer(20127));
        }
        "utf-16" | "utf16" | "unicode" => {
            fields.insert("name".to_string(), Value::String("UTF-16".to_string()));
            fields.insert("codepage".to_string(), Value::Integer(1200));
        }
        "windows-1252" | "cp1252" => {
            fields.insert("name".to_string(), Value::String("Windows-1252".to_string()));
            fields.insert("codepage".to_string(), Value::Integer(1252));
        }
        "iso-8859-1" | "latin1" => {
            fields.insert("name".to_string(), Value::String("ISO-8859-1".to_string()));
            fields.insert("codepage".to_string(), Value::Integer(28591));
        }
        _ => {
            return Err(RuntimeError::Custom(format!("Unsupported encoding: {}", encoding_name)));
        }
    }
    
    fields.insert("__encoding_type".to_string(), Value::String(encoding_name.clone()));
    
    Ok(Value::Object(Rc::new(RefCell::new(ObjectData {
        class_name: "Encoding".to_string(),
        fields,
    }))))
}

/// Encoding.Convert(srcEncoding, dstEncoding, bytes) - Convert between encodings
pub fn encoding_convert_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 3 {
        return Err(RuntimeError::Custom("Encoding.Convert requires source encoding, destination encoding, and bytes".to_string()));
    }
    
    // Get encoding types from the encoding objects
    let src_encoding = if let Value::Object(obj) = &args[0] {
        obj.borrow().fields.get("__encoding_type")
            .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
            .unwrap_or_else(|| "utf-8".to_string())
    } else {
        "utf-8".to_string()
    };
    
    let dst_encoding = if let Value::Object(obj) = &args[1] {
        obj.borrow().fields.get("__encoding_type")
            .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
            .unwrap_or_else(|| "utf-8".to_string())
    } else {
        "utf-8".to_string()
    };
    
    if let Value::Array(arr) = &args[2] {
        let bytes: Result<Vec<u8>, _> = arr.iter().map(|v| {
            v.as_integer().map(|i| i as u8)
        }).collect();
        
        match bytes {
            Ok(src_bytes) => {
                // Decode from source encoding
                let text = match src_encoding.as_str() {
                    "utf-8" | "utf8" => String::from_utf8_lossy(&src_bytes).to_string(),
                    "ascii" | "us-ascii" => String::from_utf8_lossy(&src_bytes).to_string(),
                    "utf-16" | "utf16" | "unicode" => {
                        let u16_vec: Vec<u16> = src_bytes.chunks_exact(2)
                            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                            .collect();
                        String::from_utf16_lossy(&u16_vec)
                    }
                    "windows-1252" | "cp1252" | "iso-8859-1" | "latin1" => {
                        // For Latin-1/Windows-1252, just interpret bytes as Unicode code points
                        src_bytes.iter().map(|&b| b as char).collect()
                    }
                    _ => String::from_utf8_lossy(&src_bytes).to_string(),
                };
                
                // Encode to destination encoding
                let dst_bytes: Vec<Value> = match dst_encoding.as_str() {
                    "utf-8" | "utf8" => text.as_bytes().iter().map(|&b| Value::Integer(b as i32)).collect(),
                    "ascii" | "us-ascii" => text.as_bytes().iter().map(|&b| Value::Integer(b as i32)).collect(),
                    "utf-16" | "utf16" | "unicode" => {
                        text.encode_utf16()
                            .flat_map(|u| vec![(u & 0xFF) as u8, ((u >> 8) & 0xFF) as u8])
                            .map(|b| Value::Integer(b as i32))
                            .collect()
                    }
                    "windows-1252" | "cp1252" | "iso-8859-1" | "latin1" => {
                        text.chars()
                            .map(|c| Value::Integer((c as u32).min(255) as i32))
                            .collect()
                    }
                    _ => text.as_bytes().iter().map(|&b| Value::Integer(b as i32)).collect(),
                };
                
                Ok(Value::Array(dst_bytes))
            }
            Err(e) => Err(e),
        }
    } else {
        Err(RuntimeError::Custom("Convert requires a byte array".to_string()))
    }
}

// ═══════════════════════════════════════════════════════════════════════════
// Regex Implementation
// ═══════════════════════════════════════════════════════════════════════════

/// Regex.IsMatch(input, pattern) - Test if pattern matches
pub fn regex_ismatch_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("Regex.IsMatch requires input and pattern arguments".to_string()));
    }
    
    let input = args[0].as_string();
    let pattern = args[1].as_string();
    
    match regex::Regex::new(&pattern) {
        Ok(re) => Ok(Value::Boolean(re.is_match(&input))),
        Err(e) => Err(RuntimeError::Custom(format!("Invalid regex pattern: {}", e))),
    }
}

/// Regex.Match(input, pattern) - Find first match
pub fn regex_match_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("Regex.Match requires input and pattern arguments".to_string()));
    }
    
    let input = args[0].as_string();
    let pattern = args[1].as_string();
    
    match regex::Regex::new(&pattern) {
        Ok(re) => {
            use crate::value::ObjectData;
            if let Some(m) = re.find(&input) {
                let mut fields = std::collections::HashMap::new();
                fields.insert("value".to_string(), Value::String(m.as_str().to_string()));
                fields.insert("index".to_string(), Value::Integer(m.start() as i32));
                fields.insert("length".to_string(), Value::Integer(m.len() as i32));
                fields.insert("success".to_string(), Value::Boolean(true));
                Ok(Value::Object(Rc::new(RefCell::new(ObjectData { class_name: "Match".to_string(), fields }))))
            } else {
                let mut fields = std::collections::HashMap::new();
                fields.insert("success".to_string(), Value::Boolean(false));
                Ok(Value::Object(Rc::new(RefCell::new(ObjectData { class_name: "Match".to_string(), fields }))))
            }
        },
        Err(e) => Err(RuntimeError::Custom(format!("Invalid regex pattern: {}", e))),
    }
}

/// Regex.Matches(input, pattern) - Find all matches
pub fn regex_matches_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("Regex.Matches requires input and pattern arguments".to_string()));
    }
    
    let input = args[0].as_string();
    let pattern = args[1].as_string();
    
    match regex::Regex::new(&pattern) {
        Ok(re) => {
            use crate::value::ObjectData;
            let matches: Vec<Value> = re.find_iter(&input).map(|m| {
                let mut fields = std::collections::HashMap::new();
                fields.insert("value".to_string(), Value::String(m.as_str().to_string()));
                fields.insert("index".to_string(), Value::Integer(m.start() as i32));
                fields.insert("length".to_string(), Value::Integer(m.len() as i32));
                Value::Object(Rc::new(RefCell::new(ObjectData { class_name: "Match".to_string(), fields })))
            }).collect();
            Ok(Value::Array(matches))
        },
        Err(e) => Err(RuntimeError::Custom(format!("Invalid regex pattern: {}", e))),
    }
}

/// Regex.Replace(input, pattern, replacement) - Replace matches
pub fn regex_replace_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 3 {
        return Err(RuntimeError::Custom("Regex.Replace requires input, pattern, and replacement arguments".to_string()));
    }
    
    let input = args[0].as_string();
    let pattern = args[1].as_string();
    let replacement = args[2].as_string();
    
    match regex::Regex::new(&pattern) {
        Ok(re) => Ok(Value::String(re.replace_all(&input, replacement.as_str()).to_string())),
        Err(e) => Err(RuntimeError::Custom(format!("Invalid regex pattern: {}", e))),
    }
}

/// Regex.Split(input, pattern) - Split string by pattern
pub fn regex_split_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("Regex.Split requires input and pattern arguments".to_string()));
    }
    
    let input = args[0].as_string();
    let pattern = args[1].as_string();
    
    match regex::Regex::new(&pattern) {
        Ok(re) => {
            let parts: Vec<Value> = re.split(&input)
                .map(|s| Value::String(s.to_string()))
                .collect();
            Ok(Value::Array(parts))
        },
        Err(e) => Err(RuntimeError::Custom(format!("Invalid regex pattern: {}", e))),
    }
}

// ═══════════════════════════════════════════════════════════════════════════
// JSON Implementation
// ═══════════════════════════════════════════════════════════════════════════

fn value_to_json(value: &Value) -> serde_json::Value {
    match value {
        Value::Nothing => serde_json::Value::Null,
        Value::Boolean(b) => serde_json::Value::Bool(*b),
        Value::Integer(i) => serde_json::Value::Number((*i).into()),
        Value::Long(l) => serde_json::Value::Number((*l).into()),
        Value::Double(d) => serde_json::Number::from_f64(*d)
            .map(serde_json::Value::Number)
            .unwrap_or(serde_json::Value::Null),
        Value::String(s) => serde_json::Value::String(s.clone()),
        Value::Array(arr) => {
            serde_json::Value::Array(arr.iter().map(value_to_json).collect())
        }
        Value::Object(obj) => {
            let obj_data = obj.borrow();
            let mut map = serde_json::Map::new();
            for (k, v) in obj_data.fields.iter() {
                if !k.starts_with("__") {  // Skip internal fields
                    map.insert(k.clone(), value_to_json(v));
                }
            }
            serde_json::Value::Object(map)
        }
        _ => serde_json::Value::String(format!("{:?}", value)),
    }
}

fn json_to_value(json: &serde_json::Value) -> Value {
    match json {
        serde_json::Value::Null => Value::Nothing,
        serde_json::Value::Bool(b) => Value::Boolean(*b),
        serde_json::Value::Number(n) => {
            if let Some(i) = n.as_i64() {
                if i >= i32::MIN as i64 && i <= i32::MAX as i64 {
                    Value::Integer(i as i32)
                } else {
                    Value::Long(i)
                }
            } else if let Some(f) = n.as_f64() {
                Value::Double(f)
            } else {
                Value::Nothing
            }
        }
        serde_json::Value::String(s) => Value::String(s.clone()),
        serde_json::Value::Array(arr) => {
            Value::Array(arr.iter().map(json_to_value).collect())
        }
        serde_json::Value::Object(obj) => {
            use crate::value::ObjectData;
            let mut fields = std::collections::HashMap::new();
            for (k, v) in obj.iter() {
                fields.insert(k.clone(), json_to_value(v));
            }
            Value::Object(Rc::new(RefCell::new(ObjectData { class_name: "JsonObject".to_string(), fields })))
        }
    }
}

/// JsonSerializer.Serialize(object) - Convert object to JSON string
pub fn json_serialize_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("JsonSerializer.Serialize requires an object argument".to_string()));
    }
    
    let json_value = value_to_json(&args[0]);
    match serde_json::to_string_pretty(&json_value) {
        Ok(s) => Ok(Value::String(s)),
        Err(e) => Err(RuntimeError::Custom(format!("JSON serialization error: {}", e))),
    }
}

/// JsonSerializer.Deserialize(json) - Parse JSON string to object
pub fn json_deserialize_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("JsonSerializer.Deserialize requires a JSON string argument".to_string()));
    }
    
    let json_str = args[0].as_string();
    match serde_json::from_str::<serde_json::Value>(&json_str) {
        Ok(json) => Ok(json_to_value(&json)),
        Err(e) => Err(RuntimeError::Custom(format!("JSON deserialization error: {}", e))),
    }
}

// ═══════════════════════════════════════════════════════════════════════════
// XML Implementation
// ═══════════════════════════════════════════════════════════════════════════

use quick_xml::events::Event;
use quick_xml::Reader;
use quick_xml::Writer;
use std::io::Cursor;

fn xml_to_value(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Value, RuntimeError> {
    let mut map = std::collections::HashMap::new();
    let mut current_text = String::new();
    let mut children = Vec::new();
    
    loop {
        match reader.read_event_into(buf) {
            Ok(Event::Start(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                
                // Parse attributes
                use crate::value::ObjectData;
                let mut attrs = std::collections::HashMap::new();
                for attr in e.attributes() {
                    if let Ok(a) = attr {
                        let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
                        let value = String::from_utf8_lossy(&a.value).to_string();
                        attrs.insert(key, Value::String(value));
                    }
                }
                if !attrs.is_empty() {
                    map.insert("@attributes".to_string(), Value::Object(Rc::new(RefCell::new(ObjectData { class_name: "Attributes".to_string(), fields: attrs }))));
                }
                
                // Recursively parse child element
                let child = xml_to_value(reader, buf)?;
                children.push((name, child));
            }
            Ok(Event::Text(e)) => {
                current_text.push_str(&e.unescape().unwrap_or_default());
            }
            Ok(Event::End(_)) => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(RuntimeError::Custom(format!("XML parse error: {}", e))),
            _ => {}
        }
        buf.clear();
    }
    
    // Build result
    if !current_text.trim().is_empty() && children.is_empty() {
        return Ok(Value::String(current_text.trim().to_string()));
    }
    
    for (name, child) in children {
        if let Some(existing) = map.get_mut(&name) {
            // Convert to array if multiple children with same name
            match existing {
                Value::Array(arr) => arr.push(child),
                _ => {
                    let old = existing.clone();
                    *existing = Value::Array(vec![old, child]);
                }
            }
        } else {
            map.insert(name, child);
        }
    }
    
    if !current_text.trim().is_empty() {
        map.insert("@text".to_string(), Value::String(current_text.trim().to_string()));
    }
    
    use crate::value::ObjectData;
    Ok(Value::Object(Rc::new(RefCell::new(ObjectData { class_name: "XmlElement".to_string(), fields: map }))))
}

/// XDocument.Parse(xml) - Parse XML string
pub fn xml_parse_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("XDocument.Parse requires an XML string argument".to_string()));
    }
    
    let xml_str = args[0].as_string();
    let mut reader = Reader::from_str(&xml_str);
    reader.trim_text(true);
    let mut buf = Vec::new();
    
    let mut root = std::collections::HashMap::new();
    
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                let child = xml_to_value(&mut reader, &mut buf)?;
                root.insert(name, child);
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(RuntimeError::Custom(format!("XML parse error: {}", e))),
            _ => {}
        }
        buf.clear();
    }
    
    use crate::value::ObjectData;
    Ok(Value::Object(Rc::new(RefCell::new(ObjectData { class_name: "XmlDocument".to_string(), fields: root }))))
}

fn value_to_xml(writer: &mut Writer<Cursor<Vec<u8>>>, name: &str, value: &Value) -> Result<(), RuntimeError> {
    use quick_xml::events::{BytesStart, BytesText, BytesEnd};
    
    match value {
        Value::String(s) => {
            writer.write_event(Event::Start(BytesStart::new(name)))
                .map_err(|e| RuntimeError::Custom(format!("XML write error: {}", e)))?;
            writer.write_event(Event::Text(BytesText::new(s)))
                .map_err(|e| RuntimeError::Custom(format!("XML write error: {}", e)))?;
            writer.write_event(Event::End(BytesEnd::new(name)))
                .map_err(|e| RuntimeError::Custom(format!("XML write error: {}", e)))?;
        }
        Value::Object(obj) => {
            writer.write_event(Event::Start(BytesStart::new(name)))
                .map_err(|e| RuntimeError::Custom(format!("XML write error: {}", e)))?;
            
            let obj_data = obj.borrow();
            for (k, v) in obj_data.fields.iter() {
                if !k.starts_with("@") && !k.starts_with("__") {
                    value_to_xml(writer, k, v)?;
                }
            }
            
            writer.write_event(Event::End(BytesEnd::new(name)))
                .map_err(|e| RuntimeError::Custom(format!("XML write error: {}", e)))?;
        }
        Value::Array(arr) => {
            for item in arr {
                value_to_xml(writer, name, item)?;
            }
        }
        _ => {
            let text = match value {
                Value::Integer(i) => i.to_string(),
                Value::Long(l) => l.to_string(),
                Value::Double(d) => d.to_string(),
                Value::Boolean(b) => b.to_string(),
                _ => String::new(),
            };
            writer.write_event(Event::Start(BytesStart::new(name)))
                .map_err(|e| RuntimeError::Custom(format!("XML write error: {}", e)))?;
            writer.write_event(Event::Text(BytesText::new(&text)))
                .map_err(|e| RuntimeError::Custom(format!("XML write error: {}", e)))?;
            writer.write_event(Event::End(BytesEnd::new(name)))
                .map_err(|e| RuntimeError::Custom(format!("XML write error: {}", e)))?;
        }
    }
    
    Ok(())
}

/// XDocument.Save(object) - Convert object to XML string
pub fn xml_save_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("XDocument.Save requires an object argument".to_string()));
    }
    
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    
    if let Value::Object(obj) = &args[0] {
        let obj_data = obj.borrow();
        for (name, value) in obj_data.fields.iter() {
            if !name.starts_with("__") {
                value_to_xml(&mut writer, name, value)?;
            }
        }
    } else {
        value_to_xml(&mut writer, "root", &args[0])?;
    }
    
    let result = writer.into_inner().into_inner();
    match String::from_utf8(result) {
        Ok(s) => Ok(Value::String(s)),
        Err(e) => Err(RuntimeError::Custom(format!("XML encoding error: {}", e))),
    }
}
