//! Loader for `.vbproj` (VB.NET) project files.
//!
//! Parses the XML directly with quick-xml to find:
//! - `<Compile Include="...">` entries → .vb source files to load from disk
//! - `<StartupObject>` → entry point (form name or "Sub Main")
//! - `<AssemblyName>` → project name
//!
//! The actual .vb files (including Designer.vb partial classes) are read
//! from disk relative to the .vbproj location. The VB parser handles
//! partial classes, InitializeComponent, AddHandler etc. natively.

use std::path::Path;
use quick_xml::events::Event;
use quick_xml::Reader;
use crate::bundle::{Bundle, EntryPoint, SourceFile};

pub fn load(path: &Path) -> Result<Bundle, String> {
    let xml = std::fs::read_to_string(path)
        .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;

    let project_dir = path.parent().unwrap_or(Path::new("."));

    let mut name = String::new();
    let mut startup_object = String::new();
    let mut compile_includes: Vec<String> = Vec::new();

    // Track which XML element we're inside
    let mut in_property_group = false;
    let mut current_element = String::new();

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"PropertyGroup" => in_property_group = true,
                    b"AssemblyName" if in_property_group => current_element = "AssemblyName".into(),
                    b"StartupObject" if in_property_group => current_element = "StartupObject".into(),
                    b"Compile" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"Include" {
                                compile_includes.push(
                                    String::from_utf8_lossy(&attr.value).into_owned()
                                );
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                if e.name().as_ref() == b"Compile" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"Include" {
                            compile_includes.push(
                                String::from_utf8_lossy(&attr.value).into_owned()
                            );
                        }
                    }
                }
            }
            Ok(Event::Text(e)) => {
                let text = String::from_utf8_lossy(&e).into_owned();
                match current_element.as_str() {
                    "AssemblyName" => name = text,
                    "StartupObject" => startup_object = text,
                    _ => {}
                }
                current_element.clear();
            }
            Ok(Event::End(ref e)) => {
                if e.name().as_ref() == b"PropertyGroup" {
                    in_property_group = false;
                }
                current_element.clear();
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error in {}: {e}", path.display())),
            _ => {}
        }
    }

    // If no <Compile> entries, glob all .vb files next to the .vbproj
    if compile_includes.is_empty() {
        for entry in std::fs::read_dir(project_dir).map_err(|e| e.to_string())? {
            let entry = entry.map_err(|e| e.to_string())?;
            let p = entry.path();
            if p.extension().and_then(|e| e.to_str()) == Some("vb") {
                if let Some(fname) = p.file_name().and_then(|f| f.to_str()) {
                    compile_includes.push(fname.to_string());
                }
            }
        }
    }

    // Read each .vb file from disk
    let lang = crate::languages::find_by_extension("vb")
        .ok_or("VB language not registered")?;

    let mut sources = Vec::new();
    for include in &compile_includes {
        let file_path = project_dir.join(include);
        let code = std::fs::read_to_string(&file_path)
            .map_err(|e| format!("Cannot read {}: {}", file_path.display(), e))?;
        sources.push(SourceFile { path: file_path, code });
    }

    if sources.is_empty() {
        return Err(format!("No .vb source files found in {}", path.display()));
    }

    // Fallback name from file stem
    if name.is_empty() {
        name = path.file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_else(|| "project".into());
    }

    // Determine entry point from <StartupObject>
    let entry_point = if startup_object.eq_ignore_ascii_case("Sub Main") {
        EntryPoint::Auto
    } else if !startup_object.is_empty() {
        // "Namespace.FormName" → just the form name part
        let form_name = startup_object
            .rsplit('.')
            .next()
            .unwrap_or(&startup_object)
            .to_string();
        EntryPoint::Form(form_name)
    } else {
        EntryPoint::Auto
    };

    Ok(Bundle { name, language: lang, sources, wasm_files: vec![], entry_point })
}
