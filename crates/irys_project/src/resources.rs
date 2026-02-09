
use quick_xml::events::{Event, BytesStart, BytesEnd, BytesText};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use serde::{Deserialize, Serialize};
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::fs;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct ResourceItem {
    pub name: String,
    pub value: String,
    pub comment: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq)]
pub struct ResourceManager {
    pub resources: Vec<ResourceItem>,
    pub file_path: Option<PathBuf>,
}

impl ResourceManager {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn load_from_file<P: AsRef<Path>>(path: P) -> Result<Self, Box<dyn std::error::Error>> {
        let content = fs::read_to_string(&path)?;
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

        // Simple state machine for parsing
        // Wait. quick-xml deserialization might be easier if structure matches.
        // But .resx has <data name="..."><value>...</value></data>
        // Let's use manual event parsing for flexibility.
        
        // Structure:
        // root
        //   data name="..."
        //     value
        //     comment (optional)

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"data" => {
                    // Found <data>
                    let mut name = String::new();
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"name" {
                            name = String::from_utf8(attr.value.to_vec())?;
                        }
                    }

                    let mut value = String::new();
                    let mut comment = None;

                    // Parse children of data
                    loop {
                        match reader.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.name().as_ref() == b"value" => {
                                // Read text content of value
                                if let Ok(Event::Text(e)) = reader.read_event_into(&mut content_buf) {
                                    value = e.unescape()?.into_owned();
                                }
                                // Consume </value>
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
                        resources.push(ResourceItem { name, value, comment });
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(Box::new(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(ResourceManager {
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
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        
        let root = BytesStart::new("root");
        writer.write_event(Event::Start(root.clone()))?;

        for res in &self.resources {
            let mut data = BytesStart::new("data");
            data.push_attribute(("name", res.name.as_str()));
            data.push_attribute(("xml:space", "preserve")); // Standard resx
            writer.write_event(Event::Start(data))?;

            let value_elem = BytesStart::new("value");
            writer.write_event(Event::Start(value_elem))?;
            writer.write_event(Event::Text(BytesText::new(&res.value)))?;
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
