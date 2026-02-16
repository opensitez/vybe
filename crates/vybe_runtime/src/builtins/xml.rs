// ===== XML / LINQ-to-XML support =====
// Implements XDocument, XElement, XAttribute, XComment, XDeclaration
// with full VB.NET System.Xml.Linq-style API.

use crate::value::{Value, ObjectData, RuntimeError};
use std::rc::Rc;
use std::cell::RefCell;
use std::collections::HashMap;

// ─── Internal field names ───────────────────────────────────────────────────
// __class      : "XmlDocument" | "XmlElement" | "XmlAttribute" | "XmlComment" | "XmlDeclaration"
// __name       : element/attribute tag name (String)
// __value      : text content / attribute value (String)
// __children   : Array of child XmlElement/XmlComment objects
// __attributes : Array of XmlAttribute objects
// __declaration: optional XmlDeclaration object (XmlDocument only)

// ─── Constructors ───────────────────────────────────────────────────────────

/// New XElement(name, content...)
/// content can be: string value, XAttribute, XElement, or mixed
pub fn create_xelement(args: &[Value]) -> Value {
    let name = args.first().map(|v| v.as_string()).unwrap_or_default();
    let mut fields = HashMap::new();
    fields.insert("__name".to_string(), Value::String(name.clone()));
    fields.insert("__value".to_string(), Value::String(String::new()));
    fields.insert("__children".to_string(), Value::Array(Vec::new()));
    fields.insert("__attributes".to_string(), Value::Array(Vec::new()));

    let obj = Value::Object(Rc::new(RefCell::new(ObjectData { drawing_commands: Vec::new(),
        class_name: "XmlElement".to_string(),
        fields,
    })));

    // Process remaining args as content
    for arg in args.iter().skip(1) {
        add_content_to_element(&obj, arg);
    }

    obj
}

/// New XAttribute(name, value)
pub fn create_xattribute(args: &[Value]) -> Value {
    let name = args.first().map(|v| v.as_string()).unwrap_or_default();
    let value = args.get(1).map(|v| v.as_string()).unwrap_or_default();
    let mut fields = HashMap::new();
    fields.insert("__name".to_string(), Value::String(name));
    fields.insert("__value".to_string(), Value::String(value));
    Value::Object(Rc::new(RefCell::new(ObjectData { drawing_commands: Vec::new(),
        class_name: "XmlAttribute".to_string(),
        fields,
    })))
}

/// New XDocument(content...)
pub fn create_xdocument(args: &[Value]) -> Value {
    let mut fields = HashMap::new();
    fields.insert("__children".to_string(), Value::Array(Vec::new()));
    fields.insert("__declaration".to_string(), Value::Nothing);

    let obj = Value::Object(Rc::new(RefCell::new(ObjectData { drawing_commands: Vec::new(),
        class_name: "XmlDocument".to_string(),
        fields,
    })));

    for arg in args {
        if is_xml_class(arg, "XmlDeclaration") {
            if let Value::Object(o) = &obj {
                o.borrow_mut().fields.insert("__declaration".to_string(), arg.clone());
            }
        } else {
            add_content_to_document(&obj, arg);
        }
    }

    obj
}

/// New XComment(text)
pub fn create_xcomment(args: &[Value]) -> Value {
    let text = args.first().map(|v| v.as_string()).unwrap_or_default();
    let mut fields = HashMap::new();
    fields.insert("__value".to_string(), Value::String(text));
    Value::Object(Rc::new(RefCell::new(ObjectData { drawing_commands: Vec::new(),
        class_name: "XmlComment".to_string(),
        fields,
    })))
}

/// New XDeclaration(version, encoding, standalone)
pub fn create_xdeclaration(args: &[Value]) -> Value {
    let version = args.first().map(|v| v.as_string()).unwrap_or_else(|| "1.0".to_string());
    let encoding = args.get(1).map(|v| v.as_string()).unwrap_or_else(|| "utf-8".to_string());
    let standalone = args.get(2).map(|v| v.as_string()).unwrap_or_else(|| "yes".to_string());
    let mut fields = HashMap::new();
    fields.insert("version".to_string(), Value::String(version));
    fields.insert("encoding".to_string(), Value::String(encoding));
    fields.insert("standalone".to_string(), Value::String(standalone));
    Value::Object(Rc::new(RefCell::new(ObjectData { drawing_commands: Vec::new(),
        class_name: "XmlDeclaration".to_string(),
        fields,
    })))
}

// ─── Static methods ─────────────────────────────────────────────────────────

/// XDocument.Parse(xmlString) — parse XML string into XmlDocument
pub fn xdocument_parse(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("XDocument.Parse requires an XML string argument".to_string()));
    }
    let xml_str = args[0].as_string();
    parse_xml_string(&xml_str)
}

/// XDocument.Load(filePath) — load XML from file
pub fn xdocument_load(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("XDocument.Load requires a file path argument".to_string()));
    }
    let path = args[0].as_string();
    let content = std::fs::read_to_string(&path)
        .map_err(|e| RuntimeError::Custom(format!("Failed to load XML file '{}': {}", path, e)))?;
    parse_xml_string(&content)
}

/// XElement.Parse(xmlString) — parse XML string into XmlElement (root element only)
pub fn xelement_parse(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("XElement.Parse requires an XML string argument".to_string()));
    }
    let xml_str = args[0].as_string();
    let doc = parse_xml_string(&xml_str)?;
    // Return the root element
    xml_property_access(&doc, "root")
}

// ─── Property access ────────────────────────────────────────────────────────

/// Handle property reads on XmlDocument / XmlElement / XmlAttribute
pub fn xml_property_access(obj: &Value, prop: &str) -> Result<Value, RuntimeError> {
    let (class_name, fields) = match obj {
        Value::Object(o) => {
            let b = o.borrow();
            (b.class_name.clone(), b.fields.clone())
        }
        _ => return Err(RuntimeError::Custom("Not an XML object".to_string())),
    };

    let prop_lower = prop.to_lowercase();

    match class_name.as_str() {
        "XmlDocument" => match prop_lower.as_str() {
            "root" => {
                // Return the first XmlElement child
                if let Some(Value::Array(children)) = fields.get("__children") {
                    for child in children {
                        if is_xml_class(child, "XmlElement") {
                            return Ok(child.clone());
                        }
                    }
                }
                Ok(Value::Nothing)
            }
            "declaration" => {
                Ok(fields.get("__declaration").cloned().unwrap_or(Value::Nothing))
            }
            "firstnode" | "lastnode" => {
                if let Some(Value::Array(children)) = fields.get("__children") {
                    if prop_lower == "firstnode" {
                        return Ok(children.first().cloned().unwrap_or(Value::Nothing));
                    } else {
                        return Ok(children.last().cloned().unwrap_or(Value::Nothing));
                    }
                }
                Ok(Value::Nothing)
            }
            "haselements" => {
                if let Some(Value::Array(children)) = fields.get("__children") {
                    return Ok(Value::Boolean(!children.is_empty()));
                }
                Ok(Value::Boolean(false))
            }
            _ => Ok(fields.get(&prop_lower).cloned().unwrap_or(Value::Nothing)),
        },
        "XmlElement" => match prop_lower.as_str() {
            "name" => Ok(fields.get("__name").cloned().unwrap_or(Value::String(String::new()))),
            "localname" => Ok(fields.get("__name").cloned().unwrap_or(Value::String(String::new()))),
            "value" => {
                // Value = text content; if has children, concat all text
                let direct = fields.get("__value").map(|v| v.as_string()).unwrap_or_default();
                if !direct.is_empty() {
                    return Ok(Value::String(direct));
                }
                // Concat text of all descendant text nodes
                Ok(Value::String(collect_text_content(&fields)))
            }
            "hasattributes" => {
                if let Some(Value::Array(attrs)) = fields.get("__attributes") {
                    return Ok(Value::Boolean(!attrs.is_empty()));
                }
                Ok(Value::Boolean(false))
            }
            "haselements" => {
                if let Some(Value::Array(children)) = fields.get("__children") {
                    let has = children.iter().any(|c| is_xml_class(c, "XmlElement"));
                    return Ok(Value::Boolean(has));
                }
                Ok(Value::Boolean(false))
            }
            "firstnode" | "lastnode" => {
                if let Some(Value::Array(children)) = fields.get("__children") {
                    if prop_lower == "firstnode" {
                        return Ok(children.first().cloned().unwrap_or(Value::Nothing));
                    } else {
                        return Ok(children.last().cloned().unwrap_or(Value::Nothing));
                    }
                }
                Ok(Value::Nothing)
            }
            "isempty" => {
                let val = fields.get("__value").map(|v| v.as_string()).unwrap_or_default();
                let has_children = fields.get("__children")
                    .and_then(|v| if let Value::Array(a) = v { Some(!a.is_empty()) } else { None })
                    .unwrap_or(false);
                Ok(Value::Boolean(val.is_empty() && !has_children))
            }
            "parent" => {
                // We don't track parent references (would need weak refs), return Nothing
                Ok(Value::Nothing)
            }
            "document" => Ok(Value::Nothing),
            "nodesbefore" | "nodesafter" => Ok(Value::Array(Vec::new())),
            _ => {
                // Try to find named child element as a convenience property
                if let Some(Value::Array(children)) = fields.get("__children") {
                    for child in children {
                        if let Value::Object(co) = child {
                            let cb = co.borrow();
                            if cb.class_name == "XmlElement" {
                                if let Some(Value::String(n)) = cb.fields.get("__name") {
                                    if n.eq_ignore_ascii_case(prop) {
                                        return Ok(child.clone());
                                    }
                                }
                            }
                        }
                    }
                }
                Ok(fields.get(&prop_lower).cloned().unwrap_or(Value::Nothing))
            }
        },
        "XmlAttribute" => match prop_lower.as_str() {
            "name" => Ok(fields.get("__name").cloned().unwrap_or(Value::String(String::new()))),
            "value" => Ok(fields.get("__value").cloned().unwrap_or(Value::String(String::new()))),
            _ => Ok(Value::Nothing),
        },
        "XmlComment" => match prop_lower.as_str() {
            "value" => Ok(fields.get("__value").cloned().unwrap_or(Value::String(String::new()))),
            "nodetype" => Ok(Value::String("Comment".to_string())),
            _ => Ok(Value::Nothing),
        },
        "XmlDeclaration" => {
            Ok(fields.get(&prop_lower).cloned().unwrap_or(Value::Nothing))
        }
        _ => Ok(Value::Nothing),
    }
}

// ─── Instance methods ───────────────────────────────────────────────────────

/// Handle method calls on XmlDocument / XmlElement objects
pub fn xml_method_call(obj: &Value, method: &str, args: &[Value]) -> Result<Value, RuntimeError> {
    let method_lower = method.to_lowercase();

    let class_name = match obj {
        Value::Object(o) => o.borrow().class_name.clone(),
        _ => return Err(RuntimeError::Custom("Not an XML object".to_string())),
    };

    match method_lower.as_str() {
        // ── Query methods ──────────────────────────────────────────────
        "element" => {
            // .Element("name") → first child element with that name, or Nothing
            let name = args.first().map(|v| v.as_string()).unwrap_or_default();
            let children = get_children(obj);
            for child in &children {
                if is_xml_class(child, "XmlElement") {
                    if get_xml_name(child).eq_ignore_ascii_case(&name) {
                        return Ok(child.clone());
                    }
                }
            }
            Ok(Value::Nothing)
        }
        "elements" => {
            // .Elements() or .Elements("name") → array of child elements
            let name_filter = args.first().map(|v| v.as_string());
            let children = get_children(obj);
            let result: Vec<Value> = children.iter()
                .filter(|c| {
                    if !is_xml_class(c, "XmlElement") { return false; }
                    if let Some(ref name) = name_filter {
                        get_xml_name(c).eq_ignore_ascii_case(name)
                    } else {
                        true
                    }
                })
                .cloned()
                .collect();
            Ok(Value::Array(result))
        }
        "descendants" => {
            // .Descendants() or .Descendants("name") → all descendant elements (deep)
            let name_filter = args.first().map(|v| v.as_string());
            let mut result = Vec::new();
            collect_descendants(obj, &name_filter, &mut result);
            Ok(Value::Array(result))
        }
        "descendantsandself" => {
            let name_filter = args.first().map(|v| v.as_string());
            let mut result = Vec::new();
            // Include self if it matches
            if is_xml_class(obj, "XmlElement") {
                if let Some(ref name) = name_filter {
                    if get_xml_name(obj).eq_ignore_ascii_case(name) {
                        result.push(obj.clone());
                    }
                } else {
                    result.push(obj.clone());
                }
            }
            collect_descendants(obj, &name_filter, &mut result);
            Ok(Value::Array(result))
        }
        "attribute" => {
            // .Attribute("name") → XmlAttribute or Nothing
            let name = args.first().map(|v| v.as_string()).unwrap_or_default();
            let attrs = get_attributes(obj);
            for attr in &attrs {
                if get_xml_name(attr).eq_ignore_ascii_case(&name) {
                    return Ok(attr.clone());
                }
            }
            Ok(Value::Nothing)
        }
        "attributes" => {
            // .Attributes() or .Attributes("name") → array of attributes
            let name_filter = args.first().map(|v| v.as_string());
            let attrs = get_attributes(obj);
            if let Some(name) = name_filter {
                let result: Vec<Value> = attrs.iter()
                    .filter(|a| get_xml_name(a).eq_ignore_ascii_case(&name))
                    .cloned()
                    .collect();
                Ok(Value::Array(result))
            } else {
                Ok(Value::Array(attrs))
            }
        }
        "nodes" => {
            // .Nodes() → all child nodes (elements + comments + text)
            Ok(Value::Array(get_children(obj)))
        }

        // ── Mutation methods ───────────────────────────────────────────
        "add" => {
            // .Add(content) — add child element, attribute, or text
            for arg in args {
                if class_name == "XmlDocument" {
                    add_content_to_document(obj, arg);
                } else {
                    add_content_to_element(obj, arg);
                }
            }
            Ok(Value::Nothing)
        }
        "addfirst" => {
            // .AddFirst(content) — add child at beginning
            for arg in args.iter().rev() {
                if let Value::Object(o) = obj {
                    let mut b = o.borrow_mut();
                    if is_xml_class(arg, "XmlAttribute") {
                        if let Some(Value::Array(attrs)) = b.fields.get_mut("__attributes") {
                            attrs.insert(0, arg.clone());
                        }
                    } else {
                        if let Some(Value::Array(children)) = b.fields.get_mut("__children") {
                            children.insert(0, arg.clone());
                        }
                    }
                }
            }
            Ok(Value::Nothing)
        }
        "remove" => {
            // .Remove() — remove this element from parent (no-op without parent tracking)
            // We can't truly remove without parent refs, but at least clear it
            if let Value::Object(o) = obj {
                let mut b = o.borrow_mut();
                b.fields.insert("__children".to_string(), Value::Array(Vec::new()));
                b.fields.insert("__attributes".to_string(), Value::Array(Vec::new()));
                b.fields.insert("__value".to_string(), Value::String(String::new()));
            }
            Ok(Value::Nothing)
        }
        "removeall" => {
            // .RemoveAll() — remove all children and attributes
            if let Value::Object(o) = obj {
                let mut b = o.borrow_mut();
                b.fields.insert("__children".to_string(), Value::Array(Vec::new()));
                b.fields.insert("__attributes".to_string(), Value::Array(Vec::new()));
                b.fields.insert("__value".to_string(), Value::String(String::new()));
            }
            Ok(Value::Nothing)
        }
        "removeattributes" => {
            if let Value::Object(o) = obj {
                let mut b = o.borrow_mut();
                b.fields.insert("__attributes".to_string(), Value::Array(Vec::new()));
            }
            Ok(Value::Nothing)
        }
        "removenodes" => {
            if let Value::Object(o) = obj {
                let mut b = o.borrow_mut();
                b.fields.insert("__children".to_string(), Value::Array(Vec::new()));
                b.fields.insert("__value".to_string(), Value::String(String::new()));
            }
            Ok(Value::Nothing)
        }
        "setvalue" => {
            // .SetValue(val) — set text content
            let val = args.first().map(|v| v.as_string()).unwrap_or_default();
            if let Value::Object(o) = obj {
                let mut b = o.borrow_mut();
                b.fields.insert("__value".to_string(), Value::String(val));
                // Clear children when setting text value directly
                b.fields.insert("__children".to_string(), Value::Array(Vec::new()));
            }
            Ok(Value::Nothing)
        }
        "setattributevalue" => {
            // .SetAttributeValue(name, value) — add or update attribute; remove if value is Nothing
            let attr_name = args.first().map(|v| v.as_string()).unwrap_or_default();
            let attr_val = args.get(1).cloned().unwrap_or(Value::Nothing);

            if let Value::Object(o) = obj {
                let mut b = o.borrow_mut();
                if let Some(Value::Array(attrs)) = b.fields.get_mut("__attributes") {
                    // Remove existing attribute with same name
                    attrs.retain(|a| !get_xml_name(a).eq_ignore_ascii_case(&attr_name));

                    // Add new attribute if value is not Nothing
                    if attr_val != Value::Nothing {
                        let new_attr = create_xattribute(&[
                            Value::String(attr_name),
                            Value::String(attr_val.as_string()),
                        ]);
                        attrs.push(new_attr);
                    }
                }
            }
            Ok(Value::Nothing)
        }
        "setelementvalue" => {
            // .SetElementValue(name, value) — add/update/remove child element by name
            let elem_name = args.first().map(|v| v.as_string()).unwrap_or_default();
            let elem_val = args.get(1).cloned().unwrap_or(Value::Nothing);

            if let Value::Object(o) = obj {
                let mut b = o.borrow_mut();
                if let Some(Value::Array(children)) = b.fields.get_mut("__children") {
                    // Remove existing child with same name
                    children.retain(|c| {
                        if is_xml_class(c, "XmlElement") {
                            !get_xml_name(c).eq_ignore_ascii_case(&elem_name)
                        } else {
                            true
                        }
                    });

                    // Add new element if value is not Nothing
                    if elem_val != Value::Nothing {
                        let new_elem = create_xelement(&[
                            Value::String(elem_name),
                            Value::String(elem_val.as_string()),
                        ]);
                        children.push(new_elem);
                    }
                }
            }
            Ok(Value::Nothing)
        }
        "replacewith" => {
            // .ReplaceWith(content) — replace this element's content
            // Without parent tracking, we just replace the internal content
            if let Value::Object(o) = obj {
                let old_name = {
                    let b = o.borrow();
                    b.fields.get("__name").map(|v| v.as_string()).unwrap_or_default()
                };
                let mut b = o.borrow_mut();
                b.fields.insert("__children".to_string(), Value::Array(Vec::new()));
                b.fields.insert("__attributes".to_string(), Value::Array(Vec::new()));
                b.fields.insert("__value".to_string(), Value::String(String::new()));
                b.fields.insert("__name".to_string(), Value::String(old_name));
                drop(b);
                for arg in args {
                    add_content_to_element(obj, arg);
                }
            }
            Ok(Value::Nothing)
        }

        // ── Serialization methods ──────────────────────────────────────
        "tostring" => {
            Ok(Value::String(serialize_xml(obj, 0)))
        }
        "save" => {
            // .Save(filePath)
            let path = args.first().map(|v| v.as_string()).unwrap_or_default();
            let xml_str = serialize_xml(obj, 0);
            std::fs::write(&path, &xml_str)
                .map_err(|e| RuntimeError::Custom(format!("Failed to save XML to '{}': {}", path, e)))?;
            Ok(Value::Nothing)
        }

        _ => Err(RuntimeError::Custom(format!(
            "XML object does not have method '{}'", method
        ))),
    }
}

// ─── XML parsing ────────────────────────────────────────────────────────────

/// Extract attributes from a BytesStart into owned (String, String) pairs
fn extract_attributes(start: &quick_xml::events::BytesStart) -> Vec<(String, String)> {
    let mut result = Vec::new();
    for attr in start.attributes() {
        if let Ok(a) = attr {
            let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&a.value).to_string();
            result.push((key, value));
        }
    }
    result
}

fn parse_xml_string(xml_str: &str) -> Result<Value, RuntimeError> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml_str);
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut doc_children = Vec::new();
    let mut declaration = Value::Nothing;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Decl(decl)) => {
                let version = decl.version()
                    .map(|v| String::from_utf8_lossy(&v).to_string())
                    .unwrap_or_else(|_| "1.0".to_string());
                let encoding = decl.encoding()
                    .and_then(|r| r.ok())
                    .map(|v| String::from_utf8_lossy(&v).to_string())
                    .unwrap_or_else(|| "utf-8".to_string());
                let standalone = decl.standalone()
                    .and_then(|r| r.ok())
                    .map(|v| String::from_utf8_lossy(&v).to_string())
                    .unwrap_or_else(|| "yes".to_string());
                declaration = create_xdeclaration(&[
                    Value::String(version),
                    Value::String(encoding),
                    Value::String(standalone),
                ]);
            }
            Ok(Event::Start(e)) => {
                // Extract name + attrs into owned data BEFORE clearing buf and recursing
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                let attrs = extract_attributes(&e);
                drop(e);
                buf.clear();
                let elem = parse_element_content(name, attrs, &mut reader, &mut buf)?;
                doc_children.push(elem);
            }
            Ok(Event::Empty(e)) => {
                let elem = parse_empty_element(&e)?;
                doc_children.push(elem);
            }
            Ok(Event::Comment(e)) => {
                let text = e.unescape().unwrap_or_default().to_string();
                doc_children.push(create_xcomment(&[Value::String(text)]));
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(RuntimeError::Custom(format!("XML parse error: {}", e))),
            _ => {}
        }
        buf.clear();
    }

    let mut fields = HashMap::new();
    fields.insert("__children".to_string(), Value::Array(doc_children));
    fields.insert("__declaration".to_string(), declaration);

    Ok(Value::Object(Rc::new(RefCell::new(ObjectData { drawing_commands: Vec::new(),
        class_name: "XmlDocument".to_string(),
        fields,
    }))))
}

/// Parse the content (children) of an element after its opening tag has been consumed.
/// Takes owned name and attributes to avoid borrow conflicts with the buffer.
fn parse_element_content(
    name: String,
    attrs: Vec<(String, String)>,
    reader: &mut quick_xml::Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Value, RuntimeError> {
    use quick_xml::events::Event;

    let attr_values: Vec<Value> = attrs.into_iter()
        .map(|(k, v)| create_xattribute(&[Value::String(k), Value::String(v)]))
        .collect();

    let mut children = Vec::new();
    let mut text_content = String::new();

    loop {
        match reader.read_event_into(buf) {
            Ok(Event::Start(e)) => {
                let child_name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                let child_attrs = extract_attributes(&e);
                drop(e);
                buf.clear();
                let child = parse_element_content(child_name, child_attrs, reader, buf)?;
                children.push(child);
            }
            Ok(Event::Empty(e)) => {
                let child = parse_empty_element(&e)?;
                children.push(child);
            }
            Ok(Event::Text(e)) => {
                text_content.push_str(&e.unescape().unwrap_or_default());
            }
            Ok(Event::CData(e)) => {
                text_content.push_str(&String::from_utf8_lossy(&e));
            }
            Ok(Event::Comment(e)) => {
                let ctext = e.unescape().unwrap_or_default().to_string();
                children.push(create_xcomment(&[Value::String(ctext)]));
            }
            Ok(Event::End(_)) => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(RuntimeError::Custom(format!("XML parse error: {}", e))),
            _ => {}
        }
        buf.clear();
    }

    let mut fields = HashMap::new();
    fields.insert("__name".to_string(), Value::String(name));
    fields.insert("__value".to_string(), Value::String(text_content.trim().to_string()));
    fields.insert("__children".to_string(), Value::Array(children));
    fields.insert("__attributes".to_string(), Value::Array(attr_values));

    Ok(Value::Object(Rc::new(RefCell::new(ObjectData { drawing_commands: Vec::new(),
        class_name: "XmlElement".to_string(),
        fields,
    }))))
}

fn parse_empty_element(e: &quick_xml::events::BytesStart) -> Result<Value, RuntimeError> {
    let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
    let mut attrs = Vec::new();
    for attr in e.attributes() {
        if let Ok(a) = attr {
            let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&a.value).to_string();
            attrs.push(create_xattribute(&[Value::String(key), Value::String(value)]));
        }
    }

    let mut fields = HashMap::new();
    fields.insert("__name".to_string(), Value::String(name));
    fields.insert("__value".to_string(), Value::String(String::new()));
    fields.insert("__children".to_string(), Value::Array(Vec::new()));
    fields.insert("__attributes".to_string(), Value::Array(attrs));

    Ok(Value::Object(Rc::new(RefCell::new(ObjectData { drawing_commands: Vec::new(),
        class_name: "XmlElement".to_string(),
        fields,
    }))))
}

// ─── XML serialization ─────────────────────────────────────────────────────

fn serialize_xml(val: &Value, indent: usize) -> String {
    let class_name = match val {
        Value::Object(o) => o.borrow().class_name.clone(),
        _ => return val.as_string(),
    };

    let pad = "  ".repeat(indent);

    match class_name.as_str() {
        "XmlDocument" => {
            let mut parts = Vec::new();
            if let Value::Object(o) = val {
                let b = o.borrow();
                // XML declaration
                if let Some(decl) = b.fields.get("__declaration") {
                    if *decl != Value::Nothing {
                        parts.push(serialize_xml(decl, 0));
                    }
                }
                // Children
                if let Some(Value::Array(children)) = b.fields.get("__children") {
                    for child in children {
                        parts.push(serialize_xml(child, 0));
                    }
                }
            }
            parts.join("\n")
        }
        "XmlElement" => {
            if let Value::Object(o) = val {
                let b = o.borrow();
                let name = b.fields.get("__name").map(|v| v.as_string()).unwrap_or_default();
                let text = b.fields.get("__value").map(|v| v.as_string()).unwrap_or_default();

                // Build attributes string
                let mut attr_str = String::new();
                if let Some(Value::Array(attrs)) = b.fields.get("__attributes") {
                    for attr in attrs {
                        if let Value::Object(ao) = attr {
                            let ab = ao.borrow();
                            let aname = ab.fields.get("__name").map(|v| v.as_string()).unwrap_or_default();
                            let aval = ab.fields.get("__value").map(|v| v.as_string()).unwrap_or_default();
                            attr_str.push_str(&format!(" {}=\"{}\"", aname, escape_xml_attr(&aval)));
                        }
                    }
                }

                let children = b.fields.get("__children")
                    .and_then(|v| if let Value::Array(a) = v { Some(a.clone()) } else { None })
                    .unwrap_or_default();

                if children.is_empty() && text.is_empty() {
                    // Self-closing
                    format!("{}<{}{} />", pad, name, attr_str)
                } else if children.is_empty() {
                    // Text-only element
                    format!("{}<{}{}>{}</{}>", pad, name, attr_str, escape_xml_text(&text), name)
                } else {
                    // Element with children
                    let mut parts = Vec::new();
                    parts.push(format!("{}<{}{}>", pad, name, attr_str));
                    if !text.is_empty() {
                        parts.push(format!("{}  {}", pad, escape_xml_text(&text)));
                    }
                    for child in &children {
                        parts.push(serialize_xml(child, indent + 1));
                    }
                    parts.push(format!("{}</{}>", pad, name));
                    parts.join("\n")
                }
            } else {
                String::new()
            }
        }
        "XmlComment" => {
            if let Value::Object(o) = val {
                let b = o.borrow();
                let text = b.fields.get("__value").map(|v| v.as_string()).unwrap_or_default();
                format!("{}<!--{}-->", pad, text)
            } else {
                String::new()
            }
        }
        "XmlDeclaration" => {
            if let Value::Object(o) = val {
                let b = o.borrow();
                let version = b.fields.get("version").map(|v| v.as_string()).unwrap_or_else(|| "1.0".to_string());
                let encoding = b.fields.get("encoding").map(|v| v.as_string()).unwrap_or_else(|| "utf-8".to_string());
                let standalone = b.fields.get("standalone").map(|v| v.as_string()).unwrap_or_else(|| "yes".to_string());
                format!("<?xml version=\"{}\" encoding=\"{}\" standalone=\"{}\"?>", version, encoding, standalone)
            } else {
                String::new()
            }
        }
        "XmlAttribute" => {
            if let Value::Object(o) = val {
                let b = o.borrow();
                let name = b.fields.get("__name").map(|v| v.as_string()).unwrap_or_default();
                let value = b.fields.get("__value").map(|v| v.as_string()).unwrap_or_default();
                format!("{}=\"{}\"", name, escape_xml_attr(&value))
            } else {
                String::new()
            }
        }
        _ => val.as_string(),
    }
}

// ─── Helpers ────────────────────────────────────────────────────────────────

fn is_xml_class(val: &Value, class: &str) -> bool {
    if let Value::Object(o) = val {
        o.borrow().class_name == class
    } else {
        false
    }
}

fn get_xml_name(val: &Value) -> String {
    if let Value::Object(o) = val {
        o.borrow().fields.get("__name").map(|v| v.as_string()).unwrap_or_default()
    } else {
        String::new()
    }
}

fn get_children(val: &Value) -> Vec<Value> {
    if let Value::Object(o) = val {
        let b = o.borrow();
        if let Some(Value::Array(children)) = b.fields.get("__children") {
            return children.clone();
        }
    }
    Vec::new()
}

fn get_attributes(val: &Value) -> Vec<Value> {
    if let Value::Object(o) = val {
        let b = o.borrow();
        if let Some(Value::Array(attrs)) = b.fields.get("__attributes") {
            return attrs.clone();
        }
    }
    Vec::new()
}

fn collect_descendants(val: &Value, name_filter: &Option<String>, result: &mut Vec<Value>) {
    let children = get_children(val);
    for child in &children {
        if is_xml_class(child, "XmlElement") {
            if let Some(name) = name_filter {
                if get_xml_name(child).eq_ignore_ascii_case(name) {
                    result.push(child.clone());
                }
            } else {
                result.push(child.clone());
            }
            collect_descendants(child, name_filter, result);
        }
    }
}

fn collect_text_content(fields: &HashMap<String, Value>) -> String {
    let mut text = String::new();
    if let Some(Value::String(s)) = fields.get("__value") {
        text.push_str(s);
    }
    if let Some(Value::Array(children)) = fields.get("__children") {
        for child in children {
            if let Value::Object(o) = child {
                let b = o.borrow();
                if b.class_name == "XmlElement" {
                    text.push_str(&collect_text_content(&b.fields));
                }
            }
        }
    }
    text
}

fn add_content_to_element(elem: &Value, content: &Value) {
    if let Value::Object(o) = elem {
        let mut b = o.borrow_mut();
        if is_xml_class(content, "XmlAttribute") {
            if let Some(Value::Array(attrs)) = b.fields.get_mut("__attributes") {
                attrs.push(content.clone());
            }
        } else if is_xml_class(content, "XmlElement") || is_xml_class(content, "XmlComment") {
            if let Some(Value::Array(children)) = b.fields.get_mut("__children") {
                children.push(content.clone());
            }
        } else {
            // String or other value → set as text
            let text = content.as_string();
            b.fields.insert("__value".to_string(), Value::String(text));
        }
    }
}

fn add_content_to_document(doc: &Value, content: &Value) {
    if let Value::Object(o) = doc {
        let mut b = o.borrow_mut();
        if let Some(Value::Array(children)) = b.fields.get_mut("__children") {
            children.push(content.clone());
        }
    }
}

fn escape_xml_text(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
}

fn escape_xml_attr(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
     .replace('"', "&quot;")
}

/// Check if a Value is an XML type that this module handles
pub fn is_xml_object(val: &Value) -> bool {
    if let Value::Object(o) = val {
        let cn = o.borrow().class_name.clone();
        matches!(cn.as_str(), "XmlDocument" | "XmlElement" | "XmlAttribute" | "XmlComment" | "XmlDeclaration")
    } else {
        false
    }
}
