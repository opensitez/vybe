//! System.Xml.Linq — XDocument, XElement (simplified)
//! Ported from vybe_runtime/src/builtins/xml.rs

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // XDocument.Parse(xmlString) → object tree
    vm.register_host_fn("vybe:xml", "parse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let xml = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        parse_simple_xml(&xml)
    }));

    // XDocument.Load(path) → object tree
    vm.register_host_fn("vybe:xml", "load", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let path = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        match std::fs::read_to_string(&path) {
            Ok(xml) => parse_simple_xml(&xml),
            Err(_) => Value::Null,
        }
    }));

    // Save XML object to string
    vm.register_host_fn("vybe:xml", "toString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            return Value::String(Arc::from(xml_to_string(obj).as_str()));
        }
        Value::String(Arc::from(""))
    }));
}

/// Very simplified XML parser — creates nested objects with tag, value, children, attributes
fn parse_simple_xml(xml: &str) -> Value {
    let xml = xml.trim();
    if xml.is_empty() { return Value::Null; }

    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("XDocument")));

    // Find root element
    if let Some(start) = xml.find('<') {
        if let Some(end) = xml[start+1..].find('>') {
            let tag_content = &xml[start+1..start+1+end];
            let tag_name = tag_content.split_whitespace().next().unwrap_or("");
            if !tag_name.starts_with('?') && !tag_name.starts_with('!') {
                obj.properties.insert("root".into(), parse_element(xml));
            }
        }
    }
    obj.properties.insert("__raw".into(), Value::String(Arc::from(xml)));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn parse_element(xml: &str) -> Value {
    let xml = xml.trim();
    if xml.is_empty() { return Value::Null; }

    // Find opening tag
    let start = match xml.find('<') {
        Some(i) => i,
        None => return Value::String(Arc::from(xml)),
    };
    let end = match xml[start+1..].find('>') {
        Some(i) => start + 1 + i,
        None => return Value::String(Arc::from(xml)),
    };

    let tag_content = &xml[start+1..end];
    let self_closing = tag_content.ends_with('/');
    let tag_content = tag_content.trim_end_matches('/');
    let parts: Vec<&str> = tag_content.split_whitespace().collect();
    let tag_name = parts.first().unwrap_or(&"");

    if tag_name.starts_with('?') || tag_name.starts_with('!') {
        // Skip processing instructions / comments
        return Value::Null;
    }

    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("XElement")));
    obj.properties.insert("name".into(), Value::String(Arc::from(*tag_name)));

    // Parse attributes
    for part in parts.iter().skip(1) {
        if let Some(eq_pos) = part.find('=') {
            let key = &part[..eq_pos];
            let val = part[eq_pos+1..].trim_matches('"').trim_matches('\'');
            obj.properties.insert(key.to_string(), Value::String(Arc::from(val)));
        }
    }

    if self_closing {
        obj.properties.insert("value".into(), Value::String(Arc::from("")));
    } else {
        // Find closing tag
        let close_tag = format!("</{}>", tag_name);
        if let Some(close_pos) = xml.rfind(&close_tag) {
            let inner = &xml[end+1..close_pos];
            let trimmed = inner.trim();
            if trimmed.contains('<') {
                // Has child elements
                let mut children = Vec::new();
                // Simplified: split by top-level elements
                let child = parse_element(trimmed);
                if !matches!(child, Value::Null) {
                    children.push(child);
                }
                obj.properties.insert("elements".into(),
                    Value::Object(Arc::new(Mutex::new(Object::new_array(children)))));
            } else {
                obj.properties.insert("value".into(), Value::String(Arc::from(trimmed)));
            }
        }
    }

    Value::Object(Arc::new(Mutex::new(obj)))
}

fn xml_to_string(obj: &Arc<Mutex<Object>>) -> String {
    let o = obj.lock().unwrap();
    let name = o.properties.get("name").map(|v| format!("{}", v)).unwrap_or_default();
    let value = o.properties.get("value").map(|v| format!("{}", v)).unwrap_or_default();
    if name.is_empty() {
        return o.properties.get("__raw").map(|v| format!("{}", v)).unwrap_or_default();
    }
    if value.is_empty() {
        format!("<{}/>", name)
    } else {
        format!("<{}>{}</{}>", name, value, name)
    }
}
