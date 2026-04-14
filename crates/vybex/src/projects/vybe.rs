//! Loader for `.vybe` project files (XML format).
//!
//! ```xml
//! <Project Name="MyApp" Language="javascript">
//!   <EntryPoint Type="Auto"/>
//!   <Sources>
//!     <File Path="src/main.js"/>
//!     <File Path="src/utils.js"/>
//!   </Sources>
//! </Project>
//! ```

use std::path::Path;
use quick_xml::events::Event;
use quick_xml::Reader;
use crate::bundle::{Bundle, EntryPoint, SourceFile};

pub fn load(path: &Path) -> Result<Bundle, String> {
    let xml = std::fs::read_to_string(path)
        .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;

    let project_dir = path.parent().unwrap_or(Path::new("."));

    let mut name = String::new();
    let mut language = String::new();
    let mut entry_point = EntryPoint::Auto;
    let mut source_paths: Vec<String> = Vec::new();

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match e.name().as_ref() {
                    b"Project" => {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Name" => name = String::from_utf8_lossy(&attr.value).into(),
                                b"Language" => language = String::from_utf8_lossy(&attr.value).into(),
                                _ => {}
                            }
                        }
                    }
                    b"EntryPoint" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"Type" {
                                let val = String::from_utf8_lossy(&attr.value);
                                entry_point = match val.to_lowercase().as_str() {
                                    "auto" => EntryPoint::Auto,
                                    _ => EntryPoint::Form(val.into_owned()),
                                };
                            }
                        }
                    }
                    b"File" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"Path" {
                                source_paths.push(String::from_utf8_lossy(&attr.value).into());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    if language.is_empty() {
        return Err("Missing Language attribute on <Project>".into());
    }
    if name.is_empty() {
        name = path.file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_else(|| "project".into());
    }

    let lang = crate::languages::find_by_extension(&language).ok_or_else(|| {
        format!("Unknown language '{}' in project file", language)
    })?;

    let mut sources = Vec::new();
    for rel_path in &source_paths {
        let full_path = project_dir.join(rel_path);
        let code = std::fs::read_to_string(&full_path)
            .map_err(|e| format!("Cannot read source '{}': {}", rel_path, e))?;
        sources.push(SourceFile { path: full_path, code });
    }

    if sources.is_empty() {
        return Err("Project has no source files".into());
    }

    Ok(Bundle { name, language: lang, sources, entry_point })
}
