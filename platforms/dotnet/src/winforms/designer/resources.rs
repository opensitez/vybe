use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use serde::{Deserialize, Serialize};
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

/// The category/type of a resource, matching Visual Studio's resource editor tabs.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq, Hash)]
pub enum ResourceType {
    String,
    Image,
    Icon,
    Audio,
    File,
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
    pub value: String,
    pub comment: Option<String>,
    #[serde(default)]
    pub resource_type: ResourceType,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,
}

impl ResourceItem {
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

    pub fn new_file(
        name: impl Into<String>,
        file_path: impl Into<String>,
        resource_type: ResourceType,
    ) -> Self {
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
    if lower.ends_with(".png") {
        Some("image/png".into())
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        Some("image/jpeg".into())
    } else if lower.ends_with(".gif") {
        Some("image/gif".into())
    } else if lower.ends_with(".bmp") {
        Some("image/bmp".into())
    } else {
        None
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq)]
pub struct ResourceManager {
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
        let content = crate::winforms::designer::encoding::read_text_file(&path)?;
        let mut manager = Self::parse_resx(&content)?;
        manager.file_path = Some(path.as_ref().to_path_buf());
        Ok(manager)
    }

    pub fn parse_resx(content: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let mut reader = Reader::from_str(content);
        reader.config_mut().trim_text(true);

        let mut resources = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"data" => {
                    let mut name = String::new();
                    let mut type_attr = String::new();
                    let mut mimetype_attr = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => name = String::from_utf8_lossy(&attr.value).into_owned(),
                            b"type" => {
                                type_attr = String::from_utf8_lossy(&attr.value).into_owned()
                            }
                            b"mimetype" => {
                                mimetype_attr = String::from_utf8_lossy(&attr.value).into_owned()
                            }
                            _ => {}
                        }
                    }

                    let mut value = String::new();
                    let mut comment = None;

                    loop {
                        match reader.read_event() {
                            Ok(Event::Start(ref inner)) if inner.name().as_ref() == b"value" => {
                                if let Ok(Event::Text(t)) = reader.read_event() {
                                    value = String::from_utf8_lossy(&t).into_owned();
                                }
                                // consume until </value>
                                let _ = reader.read_to_end(inner.name().to_owned());
                            }
                            Ok(Event::Start(ref inner)) if inner.name().as_ref() == b"comment" => {
                                if let Ok(Event::Text(t)) = reader.read_event() {
                                    comment = Some(String::from_utf8_lossy(&t).into_owned());
                                }
                                let _ = reader.read_to_end(inner.name().to_owned());
                            }
                            Ok(Event::End(ref inner)) if inner.name().as_ref() == b"data" => break,
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                    }

                    if !name.is_empty() {
                        let is_file_ref = type_attr.contains("ResXFileRef");
                        let (actual_value, file_name, inferred_type, mime_type) = if is_file_ref {
                            parse_file_ref_value(&value)
                        } else {
                            let rt = detect_resource_type(&type_attr, &mimetype_attr, &value);
                            let fn_ = if matches!(
                                rt,
                                ResourceType::Image
                                    | ResourceType::Icon
                                    | ResourceType::Audio
                                    | ResourceType::File
                            ) {
                                Some(value.clone())
                            } else {
                                None
                            };
                            let mt = if !mimetype_attr.is_empty() {
                                Some(mimetype_attr.clone())
                            } else {
                                None
                            };
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

        writer.write_event(Event::Decl(quick_xml::events::BytesDecl::new(
            "1.0",
            Some("utf-8"),
            None,
        )))?;

        let root = BytesStart::new("root");
        writer.write_event(Event::Start(root.clone()))?;

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

        for res in &self.resources {
            let is_file_resource = matches!(
                res.resource_type,
                ResourceType::Image | ResourceType::Icon | ResourceType::Audio | ResourceType::File
            );

            let mut data = BytesStart::new("data");
            data.push_attribute(("name", res.name.as_str()));
            data.push_attribute(("xml:space", "preserve"));

            if is_file_resource {
                data.push_attribute(("type", "System.Resources.ResXFileRef, System.Windows.Forms"));
            } else if res.resource_type == ResourceType::Other {
                if let Some(mime) = &res.mime_type {
                    data.push_attribute(("mimetype", mime.as_str()));
                }
            }

            writer.write_event(Event::Start(data))?;

            let value_elem = BytesStart::new("value");
            writer.write_event(Event::Start(value_elem))?;

            if is_file_resource {
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
            "System.Byte[], mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089".to_string()
        }
        _ => String::new(),
    }
}

fn parse_file_ref_value(raw: &str) -> (String, Option<String>, ResourceType, Option<String>) {
    let parts: Vec<&str> = raw.splitn(2, ';').collect();
    let file_path = parts[0].trim().to_string();
    let type_part = parts.get(1).unwrap_or(&"").to_lowercase();

    let (rt, mime) = if type_part.contains("bitmap") || type_part.contains("system.drawing.image") {
        (ResourceType::Image, guess_image_mime(&file_path))
    } else if type_part.contains("icon") {
        (ResourceType::Icon, Some("image/x-icon".to_string()))
    } else if type_part.contains("memorystream") || type_part.contains("audio") {
        let lower = file_path.to_lowercase();
        if lower.ends_with(".wav") || lower.ends_with(".mp3") || lower.ends_with(".ogg") {
            (ResourceType::Audio, Some("audio/wav".to_string()))
        } else {
            (ResourceType::File, None)
        }
    } else if type_part.contains("byte[]") {
        (ResourceType::File, None)
    } else {
        let rt = detect_resource_type("", "", &file_path);
        (rt, None)
    };

    (file_path.clone(), Some(file_path), rt, mime)
}

fn detect_resource_type(type_attr: &str, mimetype_attr: &str, value: &str) -> ResourceType {
    if mimetype_attr.contains("application/x-microsoft.net.object.binary") {
        return ResourceType::Other;
    }

    let type_lower = type_attr.to_lowercase();
    if type_lower.contains("bitmap") || type_lower.contains("image") {
        return ResourceType::Image;
    }
    if type_lower.contains("icon") {
        return ResourceType::Icon;
    }

    let val_lower = value.to_lowercase();
    if val_lower.ends_with(".png")
        || val_lower.ends_with(".jpg")
        || val_lower.ends_with(".jpeg")
        || val_lower.ends_with(".bmp")
        || val_lower.ends_with(".gif")
    {
        return ResourceType::Image;
    }
    if val_lower.ends_with(".ico") {
        return ResourceType::Icon;
    }
    if val_lower.ends_with(".wav") || val_lower.ends_with(".mp3") {
        return ResourceType::Audio;
    }
    if val_lower.ends_with(".txt")
        || val_lower.ends_with(".pdf")
        || val_lower.ends_with(".xml")
        || val_lower.ends_with(".json")
        || val_lower.ends_with(".pfx")
        || val_lower.ends_with(".cer")
    {
        return ResourceType::File;
    }

    ResourceType::String
}
