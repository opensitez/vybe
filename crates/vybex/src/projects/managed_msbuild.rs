//! Shared loader for managed MSBuild project files.
//!
//! Supports:
//! - `.csproj` (C#)
//! - `.pyproj` / `.ipyproj` (IronPython/Python project systems)

use std::path::{Path, PathBuf};

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::bundle::{Bundle, EntryPoint, SourceFile};
use crate::projects::encoding::read_text_file;

pub fn load(path: &Path) -> Result<Bundle, String> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_ascii_lowercase())
        .unwrap_or_default();

    let (language_ext, source_ext) = match ext.as_str() {
        "csproj" => ("cs", "cs"),
        "pyproj" | "ipyproj" => ("py", "py"),
        _ => return Err(format!("Unsupported managed project extension: .{}", ext)),
    };

    load_with_language(path, language_ext, source_ext)
}

fn load_with_language(path: &Path, language_ext: &str, source_ext: &str) -> Result<Bundle, String> {
    let xml = read_text_file(path).map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;
    let project_dir = path.parent().unwrap_or(Path::new("."));

    let mut name = String::new();
    let mut startup_object = String::new();
    let mut compile_includes: Vec<String> = Vec::new();
    let mut in_property_group = false;
    let mut current_element = String::new();

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.name().as_ref() {
                b"PropertyGroup" => in_property_group = true,
                b"AssemblyName" if in_property_group => current_element = "AssemblyName".into(),
                b"Name" if in_property_group => current_element = "Name".into(),
                b"StartupObject" if in_property_group => current_element = "StartupObject".into(),
                b"Compile" | b"Content" => push_include(e, &mut compile_includes),
                _ => {}
            },
            Ok(Event::Empty(ref e)) => {
                if matches!(e.name().as_ref(), b"Compile" | b"Content") {
                    push_include(e, &mut compile_includes);
                }
            }
            Ok(Event::Text(e)) => {
                let text = String::from_utf8_lossy(&e).into_owned();
                match current_element.as_str() {
                    "AssemblyName" | "Name" => {
                        if name.is_empty() {
                            name = text;
                        }
                    }
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

    let lang = crate::languages::find_by_extension(language_ext)
        .ok_or_else(|| format!("Language for .{} is not registered", language_ext))?;

    let mut sources = Vec::new();

    for include in compile_includes {
        let include = include.replace('\\', "/");
        if include.contains('*') {
            continue;
        }
        if !include
            .rsplit('.')
            .next()
            .is_some_and(|e| e.eq_ignore_ascii_case(source_ext))
        {
            continue;
        }

        let file_path = project_dir.join(&include);
        if !file_path.exists() {
            continue;
        }

        let code = read_text_file(&file_path)
            .map_err(|e| format!("Cannot read {}: {}", file_path.display(), e))?;
        sources.push(SourceFile {
            path: file_path,
            code,
        });
    }

    // SDK-style projects often rely on implicit includes, so fall back to
    // scanning the project directory recursively when explicit items aren't
    // present or don't resolve to source files.
    if sources.is_empty() {
        for file_path in collect_sources_recursive(project_dir, source_ext)? {
            let code = read_text_file(&file_path)
                .map_err(|e| format!("Cannot read {}: {}", file_path.display(), e))?;
            sources.push(SourceFile {
                path: file_path,
                code,
            });
        }
    }

    if sources.is_empty() {
        return Err(format!(
            "No .{} source files found for {}",
            source_ext,
            path.display()
        ));
    }

    if name.is_empty() {
        name = path
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_else(|| "project".into());
    }

    let entry_point = parse_startup_object(&startup_object, language_ext);

    Ok(Bundle {
        name,
        language: lang,
        sources,
        wasm_files: vec![],
        entry_point,
    })
}

fn parse_startup_object(startup_object: &str, language_ext: &str) -> EntryPoint {
    let startup = startup_object.trim();
    if startup.is_empty() {
        return EntryPoint::Auto;
    }

    // Keep VB behavior parity for Sub Main literal if a managed loader ever
    // routes that path here.
    if startup.eq_ignore_ascii_case("Sub Main") {
        return EntryPoint::Auto;
    }

    // Strip namespace prefix (e.g. "CosmicArcade.Program" → "Program")
    let class_name = startup.rsplit('.').next().unwrap_or(startup).to_string();

    // C# / Python: StartupObject names the class containing the static Main()
    // entry point, not a Form. Inject `ClassName.Main()` at the script tail.
    if language_ext == "cs" || language_ext == "py" {
        EntryPoint::Method(class_name, "Main".to_string())
    } else {
        EntryPoint::Auto
    }
}

fn push_include(e: &BytesStart<'_>, includes: &mut Vec<String>) {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == b"Include" {
            includes.push(String::from_utf8_lossy(&attr.value).into_owned());
        }
    }
}

fn collect_sources_recursive(dir: &Path, source_ext: &str) -> Result<Vec<PathBuf>, String> {
    let mut out = Vec::new();
    collect_sources_recursive_inner(dir, source_ext, &mut out)?;
    out.sort();
    Ok(out)
}

fn collect_sources_recursive_inner(
    dir: &Path,
    source_ext: &str,
    out: &mut Vec<PathBuf>,
) -> Result<(), String> {
    let entries = std::fs::read_dir(dir)
        .map_err(|e| format!("Cannot read directory {}: {}", dir.display(), e))?;
    for entry in entries {
        let entry = entry.map_err(|e| e.to_string())?;
        let path = entry.path();
        if path.is_dir() {
            let dir_name = path
                .file_name()
                .and_then(|s| s.to_str())
                .unwrap_or_default()
                .to_ascii_lowercase();
            if matches!(dir_name.as_str(), "bin" | "obj" | ".git" | "target") {
                continue;
            }
            collect_sources_recursive_inner(&path, source_ext, out)?;
            continue;
        }

        if path
            .extension()
            .and_then(|e| e.to_str())
            .is_some_and(|e| e.eq_ignore_ascii_case(source_ext))
        {
            out.push(path);
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn mk_temp_dir(prefix: &str) -> PathBuf {
        let mut p = std::env::temp_dir();
        p.push(format!(
            "vybex_{}_{}_{}",
            prefix,
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .as_nanos()
        ));
        std::fs::create_dir_all(&p).expect("create temp dir");
        p
    }

    #[test]
    fn loads_csproj_with_compile_items() {
        let dir = mk_temp_dir("csproj");
        let project = dir.join("App.csproj");
        let code = dir.join("Program.cs");

        std::fs::write(
            &project,
            "<Project Sdk=\"Microsoft.NET.Sdk\">\n  <PropertyGroup><AssemblyName>App</AssemblyName></PropertyGroup>\n  <ItemGroup><Compile Include=\"Program.cs\" /></ItemGroup>\n</Project>\n",
        )
        .expect("write csproj");
        std::fs::write(&code, "class Program { static void Main() {} }").expect("write code");

        let bundle = load(&project).expect("load csproj");
        assert_eq!(bundle.name, "App");
        assert_eq!(bundle.language.name, "csharp");
        assert_eq!(bundle.sources.len(), 1);

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn loads_pyproj_as_python() {
        let dir = mk_temp_dir("pyproj");
        let project = dir.join("App.pyproj");
        let code = dir.join("main.py");

        std::fs::write(
            &project,
            "<Project ToolsVersion=\"4.0\">\n  <PropertyGroup><Name>IronApp</Name></PropertyGroup>\n  <ItemGroup><Compile Include=\"main.py\" /></ItemGroup>\n</Project>\n",
        )
        .expect("write pyproj");
        std::fs::write(&code, "print('hello')").expect("write py code");

        let bundle = load(&project).expect("load pyproj");
        assert_eq!(bundle.name, "IronApp");
        assert_eq!(bundle.language.name, "python");
        assert_eq!(bundle.sources.len(), 1);

        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn falls_back_to_recursive_scan_for_implicit_includes() {
        let dir = mk_temp_dir("implicit");
        let project = dir.join("App.csproj");
        let src_dir = dir.join("src");
        let code = src_dir.join("Program.cs");
        std::fs::create_dir_all(&src_dir).expect("mkdir src");

        std::fs::write(
            &project,
            "<Project Sdk=\"Microsoft.NET.Sdk\">\n  <PropertyGroup><AssemblyName>ImplicitApp</AssemblyName></PropertyGroup>\n</Project>\n",
        )
        .expect("write csproj");
        std::fs::write(&code, "class Program { static void Main() {} }").expect("write code");

        let bundle = load(&project).expect("load csproj implicit");
        assert_eq!(bundle.language.name, "csharp");
        assert_eq!(bundle.sources.len(), 1);
        assert!(bundle.sources[0].path.ends_with("src/Program.cs"));

        let _ = std::fs::remove_dir_all(&dir);
    }
}