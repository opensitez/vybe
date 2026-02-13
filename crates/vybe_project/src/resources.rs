
use quick_xml::events::{Event, BytesStart, BytesEnd, BytesText};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use serde::{Deserialize, Serialize};
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::fs;

/// The category/type of a resource, matching Visual Studio's resource editor tabs.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq, Hash)]
pub enum ResourceType {
    /// Plain string value (default)
    String,
    /// Image file (.png, .jpg, .bmp, .gif)
    Image,
    /// Icon file (.ico)
    Icon,
    /// Audio file (.wav)
    Audio,
    /// Arbitrary binary file (stored as base64)
    File,
    /// Other/serialized object
    Other,
}

impl Default for ResourceType {
    fn default() -> Self {
        ResourceType::String
    }
}

impl std::fmt::Display for ResourceType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ResourceType::String => write!(f, "Strings"),
            ResourceType::Image => write!(f, "Images"),
            ResourceType::Icon => write!(f, "Icons"),
            ResourceType::Audio => write!(f, "Audio"),
            ResourceType::File => write!(f, "Files"),
            ResourceType::Other => write!(f, "Other"),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct ResourceItem {
    pub name: String,
    /// For String resources: the text value.
    /// For file-based resources (Image, Icon, Audio, File): the relative file path.
    /// For binary-embedded resources: base64-encoded data.
    pub value: String,
    pub comment: Option<String>,
    /// The resource type/category
    #[serde(default)]
    pub resource_type: ResourceType,
    /// Original file name (for file-based resources)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    /// MIME type hint (e.g. "image/png", "audio/wav")
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,
}

impl ResourceItem {
    /// Create a new string resource
    pub fn new_string(name: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            value: value.into(),
            comment: None,
            resource_type: ResourceType::String,
            file_name: None,
            mime_type: None,
        }
    }

    /// Create a file-based resource (image, icon, audio, or generic file)
    pub fn new_file(name: impl Into<String>, file_path: impl Into<String>, resource_type: ResourceType) -> Self {
        let fp: String = file_path.into();
        let mime = match &resource_type {
            ResourceType::Image => guess_image_mime(&fp),
            ResourceType::Icon => Some("image/x-icon".to_string()),
            ResourceType::Audio => Some("audio/wav".to_string()),
            _ => None,
        };
        Self {
            name: name.into(),
            value: fp.clone(),
            comment: None,
            resource_type,
            file_name: Some(fp),
            mime_type: mime,
        }
    }
}

fn guess_image_mime(path: &str) -> Option<String> {
    let lower = path.to_lowercase();
    if lower.ends_with(".png") { Some("image/png".into()) }
    else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") { Some("image/jpeg".into()) }
    else if lower.ends_with(".gif") { Some("image/gif".into()) }
    else if lower.ends_with(".bmp") { Some("image/bmp".into()) }
    else { None }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq)]
pub struct ResourceManager {
    /// Display name of this resource file (e.g. "Resources", "Strings", "Images")
    #[serde(default = "default_resource_name")]
    pub name: String,
    pub resources: Vec<ResourceItem>,
    pub file_path: Option<PathBuf>,
}

fn default_resource_name() -> String {
    "Resources".to_string()
}

impl ResourceManager {
    pub fn new() -> Self {
        Self {
            name: "Resources".to_string(),
            ..Self::default()
        }
    }

    pub fn new_named(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            ..Self::default()
        }
    }

    pub fn load_from_file<P: AsRef<Path>>(path: P) -> Result<Self, Box<dyn std::error::Error>> {
        let content = crate::encoding::read_text_file(&path)?;
        let mut manager = Self::parse_resx(&content)?;
        manager.file_path = Some(path.as_ref().to_path_buf());
        Ok(manager)
    }

    pub fn parse_resx(content: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let mut reader = Reader::from_str(content);
        reader.trim_text(true);

        let mut resources = Vec::new();
        let mut buf = Vec::new();
        let mut inner_buf = Vec::new();
        let mut content_buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"data" => {
                    let mut name = String::new();
                    let mut type_attr = String::new();
                    let mut mimetype_attr = String::new();
                    for attr in e.attributes() {
                        let attr = attr?;
                        match attr.key.as_ref() {
                            b"name" => name = String::from_utf8(attr.value.to_vec())?,
                            b"type" => type_attr = String::from_utf8(attr.value.to_vec())?,
                            b"mimetype" => mimetype_attr = String::from_utf8(attr.value.to_vec())?,
                            _ => {}
                        }
                    }

                    let mut value = String::new();
                    let mut comment = None;

                    // Parse children of data
                    loop {
                        match reader.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.name().as_ref() == b"value" => {
                                if let Ok(Event::Text(e)) = reader.read_event_into(&mut content_buf) {
                                    value = e.unescape()?.into_owned();
                                }
                                reader.read_to_end(e.name().to_owned())?; 
                            }
                            Ok(Event::Start(ref e)) if e.name().as_ref() == b"comment" => {
                                if let Ok(Event::Text(e)) = reader.read_event_into(&mut content_buf) {
                                    comment = Some(e.unescape()?.into_owned());
                                }
                                reader.read_to_end(e.name().to_owned())?;
                            }
                            Ok(Event::End(ref e)) if e.name().as_ref() == b"data" => {
                                break;
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                        inner_buf.clear();
                        content_buf.clear();
                    }
                    
                    if !name.is_empty() {
                        // Handle ResXFileRef values: "path;System.Type, Assembly, ..."
                        // Real .resx encodes file refs as semicolon-delimited strings
                        let is_file_ref = type_attr.contains("ResXFileRef");
                        let (actual_value, file_name, inferred_type, mime_type) = if is_file_ref {
                            parse_file_ref_value(&value)
                        } else {
                            // Detect from type/mimetype/extension for backward compat
                            let rt = detect_resource_type(&type_attr, &mimetype_attr, &value);
                            let fn_ = if matches!(rt, ResourceType::Image | ResourceType::Icon | ResourceType::Audio | ResourceType::File) {
                                Some(value.clone())
                            } else {
                                None
                            };
                            let mt = if !mimetype_attr.is_empty() { Some(mimetype_attr.clone()) } else { None };
                            (value.clone(), fn_, rt, mt)
                        };

                        resources.push(ResourceItem {
                            name,
                            value: actual_value,
                            comment,
                            resource_type: inferred_type,
                            file_name,
                            mime_type,
                        });
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(Box::new(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(ResourceManager {
            name: "Resources".to_string(),
            resources,
            file_path: None,
        })
    }

    pub fn save(&self) -> Result<(), Box<dyn std::error::Error>> {
        if let Some(path) = &self.file_path {
            let content = self.to_resx()?;
            fs::write(path, content)?;
            Ok(())
        } else {
            Err("No file path set for resource manager".into())
        }
    }

    pub fn to_resx(&self) -> Result<String, Box<dyn std::error::Error>> {
        let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);

        // XML declaration
        writer.write_event(Event::Decl(quick_xml::events::BytesDecl::new("1.0", Some("utf-8"), None)))?;

        let root = BytesStart::new("root");
        writer.write_event(Event::Start(root.clone()))?;

        // -- resheader block (required by VS) --
        write_resheader(&mut writer, "resmimetype", "text/microsoft-resx")?;
        write_resheader(&mut writer, "version", "2.0")?;
        write_resheader(
            &mut writer,
            "reader",
            "System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
        )?;
        write_resheader(
            &mut writer,
            "writer",
            "System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
        )?;

        // -- data elements --
        for res in &self.resources {
            let is_file_resource = matches!(
                res.resource_type,
                ResourceType::Image | ResourceType::Icon | ResourceType::Audio | ResourceType::File
            );

            let mut data = BytesStart::new("data");
            data.push_attribute(("name", res.name.as_str()));
            data.push_attribute(("xml:space", "preserve"));

            if is_file_resource {
                // Real .resx uses ResXFileRef type for file-based resources
                data.push_attribute(("type", "System.Resources.ResXFileRef, System.Windows.Forms"));
            } else if res.resource_type == ResourceType::Other {
                if let Some(mime) = &res.mime_type {
                    data.push_attribute(("mimetype", mime.as_str()));
                }
            }
            // String resources: no type attribute (that's the VS default)

            writer.write_event(Event::Start(data))?;

            let value_elem = BytesStart::new("value");
            writer.write_event(Event::Start(value_elem))?;

            if is_file_resource {
                // Encode as "path;TypeName, Assembly" (VS ResXFileRef format)
                let type_ref = file_ref_type_string(&res.resource_type, &res.value);
                let ref_value = format!("{};{}", res.value, type_ref);
                writer.write_event(Event::Text(BytesText::new(&ref_value)))?;
            } else {
                writer.write_event(Event::Text(BytesText::new(&res.value)))?;
            }

            writer.write_event(Event::End(BytesEnd::new("value")))?;

            if let Some(comment) = &res.comment {
                let comment_elem = BytesStart::new("comment");
                writer.write_event(Event::Start(comment_elem))?;
                writer.write_event(Event::Text(BytesText::new(comment)))?;
                writer.write_event(Event::End(BytesEnd::new("comment")))?;
            }

            writer.write_event(Event::End(BytesEnd::new("data")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("root")))?;

        let result = String::from_utf8(writer.into_inner().into_inner())?;
        Ok(result)
    }
}

/// Write a <resheader name="..."><value>...</value></resheader> element.
fn write_resheader<W: std::io::Write>(
    writer: &mut Writer<W>,
    name: &str,
    value: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut elem = BytesStart::new("resheader");
    elem.push_attribute(("name", name));
    writer.write_event(Event::Start(elem))?;

    let val = BytesStart::new("value");
    writer.write_event(Event::Start(val))?;
    writer.write_event(Event::Text(BytesText::new(value)))?;
    writer.write_event(Event::End(BytesEnd::new("value")))?;

    writer.write_event(Event::End(BytesEnd::new("resheader")))?;
    Ok(())
}

/// Returns the assembly-qualified type string for a ResXFileRef value.
/// This is the part after the semicolon in the <value> element.
fn file_ref_type_string(resource_type: &ResourceType, file_path: &str) -> String {
    match resource_type {
        ResourceType::Image => {
            let lower = file_path.to_lowercase();
            if lower.ends_with(".png") || lower.ends_with(".jpg") || lower.ends_with(".jpeg")
                || lower.ends_with(".bmp") || lower.ends_with(".gif") || lower.ends_with(".tiff") {
                "System.Drawing.Bitmap, System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a".to_string()
            } else {
                "System.Drawing.Bitmap, System.Drawing".to_string()
            }
        }
        ResourceType::Icon => {
            "System.Drawing.Icon, System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a".to_string()
        }
        ResourceType::Audio => {
            "System.IO.MemoryStream, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089".to_string()
        }
        ResourceType::File => {
            // Generic binary file → byte array
            "System.Byte[], mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089".to_string()
        }
        _ => String::new(),
    }
}

/// Parse a ResXFileRef value string: "path;TypeName, Assembly, Version=..., ..."
/// Returns (file_path, file_name, resource_type, mime_type).
fn parse_file_ref_value(raw: &str) -> (String, Option<String>, ResourceType, Option<String>) {
    // Split on ';' — first part is the file path, rest is the .NET type
    let parts: Vec<&str> = raw.splitn(2, ';').collect();
    let file_path = parts[0].trim().to_string();
    let type_part = parts.get(1).unwrap_or(&"").to_lowercase();

    let (rt, mime) = if type_part.contains("bitmap") || type_part.contains("system.drawing.image") {
        (ResourceType::Image, guess_image_mime(&file_path))
    } else if type_part.contains("icon") {
        (ResourceType::Icon, Some("image/x-icon".to_string()))
    } else if type_part.contains("memorystream") || type_part.contains("audio") {
        // Audio files are typically stored as MemoryStream in .resx
        let lower = file_path.to_lowercase();
        if lower.ends_with(".wav") || lower.ends_with(".mp3") || lower.ends_with(".ogg") {
            (ResourceType::Audio, Some("audio/wav".to_string()))
        } else {
            (ResourceType::File, None)
        }
    } else if type_part.contains("byte[]") {
        (ResourceType::File, None)
    } else {
        // Fall back to extension-based detection
        let rt = detect_resource_type("", "", &file_path);
        (rt, None)
    };

    (file_path.clone(), Some(file_path), rt, mime)
}

/// Detect resource type from .resx XML attributes and value
fn detect_resource_type(type_attr: &str, mimetype_attr: &str, value: &str) -> ResourceType {
    // Check mimetype attribute first (binary/base64 embedded resources)
    if mimetype_attr.contains("application/x-microsoft.net.object.binary") {
        return ResourceType::Other;
    }

    // Check type attribute (e.g. "System.Drawing.Bitmap", "System.Drawing.Icon")
    let type_lower = type_attr.to_lowercase();
    if type_lower.contains("bitmap") || type_lower.contains("image") {
        return ResourceType::Image;
    }
    if type_lower.contains("icon") {
        return ResourceType::Icon;
    }

    // Check file extension in the value
    let val_lower = value.to_lowercase();
    if val_lower.ends_with(".png") || val_lower.ends_with(".jpg") || val_lower.ends_with(".jpeg")
        || val_lower.ends_with(".bmp") || val_lower.ends_with(".gif") {
        return ResourceType::Image;
    }
    if val_lower.ends_with(".ico") {
        return ResourceType::Icon;
    }
    if val_lower.ends_with(".wav") || val_lower.ends_with(".mp3") {
        return ResourceType::Audio;
    }
    if val_lower.ends_with(".txt") || val_lower.ends_with(".pdf") || val_lower.ends_with(".xml")
        || val_lower.ends_with(".json") || val_lower.ends_with(".pfx") || val_lower.ends_with(".cer") {
        return ResourceType::File;
    }

    // Default: string
    ResourceType::String
}
