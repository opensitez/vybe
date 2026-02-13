
use crate::project::{Project, FormModule};
use crate::winforms::{load_form_vb, save_form_vb};
use crate::errors::{SaveError, SaveResult};
use crate::resources::ResourceManager;
use vybe_forms::Form;
use std::path::Path;
use std::fs;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use crate::encoding::read_text_file;

pub fn save_project_auto(project: &Project, path: impl AsRef<Path>) -> SaveResult<()> {
    let path = path.as_ref();
    if let Some(ext) = path.extension() {
        if !ext.eq_ignore_ascii_case("vbproj") {
            return Err(SaveError::Parse("Only .vbproj files are supported for saving".to_string()));
        }
    }
    save_project_vbproj(project, path)
}

pub fn save_project_vbproj(project: &Project, path: impl AsRef<Path>) -> SaveResult<()> {
    let path = path.as_ref();

    let mut xml = String::new();
    xml.push_str("<Project Sdk=\"Microsoft.NET.Sdk\">\n");
    xml.push_str("  <PropertyGroup>\n");
    xml.push_str("    <OutputType>WinExe</OutputType>\n");
    xml.push_str(&format!("    <RootNamespace>{}</RootNamespace>\n", project.name));
    xml.push_str(&format!("    <AssemblyName>{}</AssemblyName>\n", project.name));
    xml.push_str("    <TargetFramework>net6.0-windows</TargetFramework>\n");
    xml.push_str("    <UseWindowsForms>true</UseWindowsForms>\n");
    match &project.startup_object {
        crate::project::StartupObject::Form(form_name) => {
            xml.push_str(&format!("    <StartupObject>{}.{}</StartupObject>\n", project.name, form_name));
        }
        crate::project::StartupObject::SubMain => {
            xml.push_str("    <StartupObject>Sub Main</StartupObject>\n");
        }
        crate::project::StartupObject::None => {
            // Don't write a StartupObject tag — loader will auto-detect
        }
    }
    xml.push_str("  </PropertyGroup>\n");
    xml.push_str("  <ItemGroup>\n");
    for form_mod in &project.forms {
        let form_name = &form_mod.form.name;
        xml.push_str(&format!("    <Compile Include=\"{}.vb\">\n", form_name));
        xml.push_str("      <SubType>Form</SubType>\n");
        xml.push_str("    </Compile>\n");
        xml.push_str(&format!("    <Compile Include=\"{}.Designer.vb\">\n", form_name));
        xml.push_str(&format!("      <DependentUpon>{}.vb</DependentUpon>\n", form_name));
        xml.push_str("    </Compile>\n");
    }
    for code_file in &project.code_files {
        xml.push_str(&format!("    <Compile Include=\"{}.vb\" />\n", code_file.name));
    }
    xml.push_str("  </ItemGroup>\n");
    xml.push_str("</Project>\n");

    fs::write(path, xml)?;

    // Save Forms
    let parent_dir = path.parent().unwrap_or(Path::new("."));
    for form_mod in &project.forms {
        let mut fm = form_mod.clone();
        fm.sync_designer_code();
        save_form_vb(&fm, parent_dir)?;
    }

    // Save code files
    for code_file in &project.code_files {
        let mod_path = parent_dir.join(format!("{}.vb", code_file.name));
        fs::write(&mod_path, &code_file.code)?;
    }

    Ok(())
}

// End of serialization.rs - Only .vbproj and .vb WinForms support

pub fn load_project_auto(path: impl AsRef<Path>) -> SaveResult<Project> {
    let path = path.as_ref();
    eprintln!("[DEBUG] load_project_auto: {:?}", path);
    if let Some(ext) = path.extension() {
        eprintln!("[DEBUG] extension: {:?}", ext);
        if ext.eq_ignore_ascii_case("vbproj") {
            return load_project_vbproj(path);
        }
    }
    Err(SaveError::Parse("Only .vbproj files are supported".to_string()))
}

pub fn load_project_vbproj(path: impl AsRef<Path>) -> SaveResult<Project> {
    let path = path.as_ref();
    let content = read_text_file(path)?;
    
    // Only support real XML .vbproj files
    let trimmed = content.trim();
    if !trimmed.starts_with('<') {
        return Err(SaveError::Parse("File is not a valid XML .vbproj".to_string()));
    }
    
    let mut reader = Reader::from_str(&content);
    reader.trim_text(true);

    let mut project_name = String::new();
    let mut startup_object = None;
    let mut form_paths = Vec::new();
    let mut module_paths = Vec::new();
    let mut resource_paths: Vec<(String, Option<String>)> = Vec::new(); // (resx_path, dependent_upon)
    let mut project_ref_paths: Vec<String> = Vec::new(); // relative paths to referenced .vbproj files

    // State tracking
    let mut current_file_path = String::new();
    let mut current_subtype = String::new();
    let mut current_dependent_upon = String::new();
    let mut capture_text = false;
    let mut current_tag = Vec::new();
    let mut in_compile = false;
    let mut in_embedded_resource = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.name().as_ref().to_vec();
                if name == b"AssemblyName" || name == b"StartupObject" || name == b"SubType" || name == b"DependentUpon" {
                    capture_text = true;
                    current_tag = name.clone();
                } else if name == b"Compile" {
                    in_compile = true;
                    current_file_path = String::new();
                    current_subtype = String::new();
                    current_dependent_upon = String::new();
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"Include" {
                                if let Ok(val) = attr.unescape_value() {
                                    current_file_path = val.into_owned();
                                }
                            }
                        }
                    }
                } else if name == b"EmbeddedResource" {
                    in_embedded_resource = true;
                    current_file_path = String::new();
                    current_dependent_upon = String::new();
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"Include" {
                                if let Ok(val) = attr.unescape_value() {
                                    current_file_path = val.into_owned();
                                }
                            }
                        }
                    }
                } else if name == b"ProjectReference" {
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"Include" {
                                if let Ok(val) = attr.unescape_value() {
                                    project_ref_paths.push(val.into_owned().replace('\\', "/"));
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                if name.as_ref() == b"Compile" {
                    let mut file_path = String::new();
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"Include" {
                                if let Ok(val) = attr.unescape_value() {
                                    file_path = val.into_owned();
                                }
                            }
                        }
                    }
                    if !file_path.is_empty() {
                         let clean_path = file_path.replace('\\', "/");
                         if !clean_path.ends_with(".Designer.vb") {
                             // Empty Compile tag usually means no SubType, so it's a module/class
                             module_paths.push(clean_path);
                         }
                    }
                } else if name.as_ref() == b"EmbeddedResource" {
                    // Self-closing <EmbeddedResource Include="Foo.resx" />
                    let mut file_path = String::new();
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"Include" {
                                if let Ok(val) = attr.unescape_value() {
                                    file_path = val.into_owned();
                                }
                            }
                        }
                    }
                    if !file_path.is_empty() {
                        let clean_path = file_path.replace('\\', "/");
                        resource_paths.push((clean_path, None));
                    }
                } else if name.as_ref() == b"ProjectReference" {
                    // Self-closing <ProjectReference Include="..." />  (SDK-style)
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"Include" {
                                if let Ok(val) = attr.unescape_value() {
                                    project_ref_paths.push(val.into_owned().replace('\\', "/"));
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if capture_text {
                    if let Ok(txt) = e.unescape() {
                        let txt = txt.into_owned();
                        if current_tag == b"AssemblyName" {
                            project_name = txt;
                        } else if current_tag == b"StartupObject" {
                            // Store the raw startup object text for now
                            if txt == "Sub Main" || txt.is_empty() {
                                // Will set to SubMain later
                                startup_object = Some("Sub Main".to_string());
                            } else if !txt.contains("My.MyApplication") {
                                startup_object = Some(txt);
                            } else {
                                // My.MyApplication — mark for Application.myapp lookup
                                startup_object = Some("__MY_APPLICATION__".to_string());
                            }
                        } else if current_tag == b"SubType" && in_compile {
                            current_subtype = txt;
                        } else if current_tag == b"DependentUpon" && (in_compile || in_embedded_resource) {
                            current_dependent_upon = txt;
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let qname = e.name();
                let name = qname.as_ref();
                if name == b"AssemblyName" || name == b"StartupObject" || name == b"SubType" || name == b"DependentUpon" {
                    capture_text = false;
                } else if name == b"Compile" {
                    in_compile = false;
                    if !current_file_path.is_empty() {
                         // Convert backslashes to forward slashes for cross-platform
                         let clean_path = current_file_path.replace('\\', "/");
                         
                         // Skip designer files
                         if !clean_path.ends_with(".Designer.vb") {
                             if current_subtype == "Form" {
                                 form_paths.push(clean_path);
                             } else {
                                 module_paths.push(clean_path);
                             }
                         }
                    }
                } else if name == b"EmbeddedResource" {
                    in_embedded_resource = false;
                    if !current_file_path.is_empty() {
                        let clean_path = current_file_path.replace('\\', "/");
                        let dep = if current_dependent_upon.is_empty() {
                            None
                        } else {
                            Some(current_dependent_upon.replace('\\', "/"))
                        };
                        resource_paths.push((clean_path, dep));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(SaveError::Parse(format!("XML error: {}", e))),
            _ => (),
        }
    }

    if project_name.is_empty() {
        project_name = path.file_stem().unwrap_or_default().to_string_lossy().to_string();
    }

    let mut project = Project::new(&project_name);
    
    // Set startup object based on what was parsed
    if let Some(ref startup_str) = startup_object {
        eprintln!("[DEBUG] Parsed startup_object string: {:?}", startup_str);
        if startup_str == "Sub Main" {
            project.startup_object = crate::project::StartupObject::SubMain;
            eprintln!("[DEBUG] Set startup_object to SubMain");
        } else {
            // Extract form name from "ProjectName.FormName" format
            let form_name = if let Some(dot_pos) = startup_str.rfind('.') {
                startup_str[dot_pos + 1..].to_string()
            } else {
                startup_str.clone()
            };
            project.startup_object = crate::project::StartupObject::Form(form_name.clone());
            project.startup_form = Some(form_name);
            eprintln!("[DEBUG] Set startup_object to Form");
        }
    } else {
        eprintln!("[DEBUG] No startup_object string parsed, setting to None");
        project.startup_object = crate::project::StartupObject::None;
    }

    let parent_dir = path.parent().unwrap_or(Path::new("."));

    // If startup is My.MyApplication, read MainForm from Application.myapp
    if startup_object.as_deref() == Some("__MY_APPLICATION__") {
        let myapp_path = parent_dir.join("My Project/Application.myapp");
        if myapp_path.exists() {
            if let Ok(myapp_content) = read_text_file(&myapp_path) {
                // Simple parse: find <MainForm>FormName</MainForm>
                if let Some(start) = myapp_content.find("<MainForm>") {
                    let after = &myapp_content[start + 10..];
                    if let Some(end) = after.find("</MainForm>") {
                        let main_form = after[..end].trim().to_string();
                        if !main_form.is_empty() {
                            eprintln!("[DEBUG] Found MainForm='{}' in Application.myapp", main_form);
                            project.startup_object = crate::project::StartupObject::Form(main_form.clone());
                            project.startup_form = Some(main_form);
                        }
                    }
                }
            }
        }
        // If still not set, default to first form
        if matches!(project.startup_object, crate::project::StartupObject::None) {
            eprintln!("[DEBUG] Application.myapp not found or no MainForm, will default to first form");
        }
    }

    // SDK-style projects have no <Compile> items — auto-discover all .vb files recursively
    if form_paths.is_empty() && module_paths.is_empty() {
        eprintln!("[DEBUG] No <Compile> items found — SDK-style project, auto-discovering .vb files");
        // Recursively collect all .vb files (like MSBuild SDK-style)
        fn collect_vb_files(dir: &Path, base: &Path, form_paths: &mut Vec<String>, module_paths: &mut Vec<String>) {
            if let Ok(entries) = fs::read_dir(dir) {
                for entry in entries.flatten() {
                    let p = entry.path();
                    if p.is_dir() {
                        // Skip common non-source directories
                        let dir_name = p.file_name().unwrap_or_default().to_string_lossy().to_lowercase();
                        if dir_name == "bin" || dir_name == "obj" || dir_name == ".git" {
                            continue;
                        }
                        collect_vb_files(&p, base, form_paths, module_paths);
                    } else if let Some(ext) = p.extension() {
                        if ext.eq_ignore_ascii_case("vb") {
                            // Use relative path from project dir
                            let rel = p.strip_prefix(base).unwrap_or(&p).to_string_lossy().to_string();
                            // Skip designer files
                            if rel.to_lowercase().ends_with(".designer.vb") {
                                continue;
                            }
                            // Check if it's a form file
                            if let Ok(raw) = read_text_file(&p) {
                                let upper = raw.to_uppercase();
                                if upper.contains("INHERITS SYSTEM.WINDOWS.FORMS.FORM")
                                    || upper.contains("INHERITS FORM")
                                {
                                    form_paths.push(rel);
                                } else {
                                    module_paths.push(rel);
                                }
                            } else {
                                module_paths.push(rel);
                            }
                        }
                    }
                }
            }
        }
        collect_vb_files(parent_dir, parent_dir, &mut form_paths, &mut module_paths);
        eprintln!("[DEBUG] Auto-discovered form_paths={:?}, module_paths={:?}", form_paths, module_paths);
    }

    eprintln!("[DEBUG] load_project_vbproj: project='{}', form_paths={:?}, module_paths={:?}, resource_paths={:?}", project_name, form_paths, module_paths, resource_paths);

    // Load Forms
    for rel_path in &form_paths {
        let form_path = parent_dir.join(&rel_path);
        eprintln!("[DEBUG] Loading form: {:?}", form_path);
        match load_form_vb(&form_path) {
            Ok(form_module) => {
                eprintln!("[DEBUG] Form loaded OK: '{}' with {} controls", form_module.form.name, form_module.form.controls.len());
                let form_name = form_module.form.name.clone();
                project.forms.push(form_module);
                
                // Update startup_object if it references this form
                if let crate::project::StartupObject::Form(ref startup_name) = project.startup_object {
                    // StartupObject might be "ProjectName.FormName" or just "FormName"
                    if startup_name.ends_with(&form_name) {
                        // Normalize to just the form name
                        project.startup_object = crate::project::StartupObject::Form(form_name.clone());
                        project.startup_form = Some(form_name);
                    }
                }
            }
            Err(e) => {
                eprintln!("[WARN] Failed to parse form {}: {} — adding as unparsed form", rel_path, e);
                // Create a fallback FormModule so the form still appears in the project explorer
                let stem = Path::new(rel_path).file_stem().unwrap_or_default().to_string_lossy().to_string();
                let user_code = read_text_file(&form_path).unwrap_or_default();
                let designer_path = form_path.with_extension("Designer.vb");
                // Also try: parent/Stem.Designer.vb
                let designer_path2 = parent_dir.join(format!("{}.Designer.vb", stem));
                let designer_code = if designer_path.exists() {
                    read_text_file(&designer_path).unwrap_or_default()
                } else if designer_path2.exists() {
                    read_text_file(&designer_path2).unwrap_or_default()
                } else {
                    String::new()
                };
                let form = Form::new(&stem);
                let fm = FormModule::new_vbnet(form, designer_code, user_code);
                let form_name = stem.clone();
                project.forms.push(fm);

                if let crate::project::StartupObject::Form(ref startup_name) = project.startup_object {
                    if startup_name.ends_with(&form_name) {
                        project.startup_object = crate::project::StartupObject::Form(form_name.clone());
                        project.startup_form = Some(form_name);
                    }
                }
            }
        }
    }

    // Load Modules
    for rel_path in module_paths {
        let mod_path = parent_dir.join(&rel_path);
        if let Ok(content) = read_text_file(&mod_path) {
            // Preserve folder prefix in name (e.g. "Extensions/GenericExtensions")
            let name = if rel_path.contains('/') {
                // Strip only the .vb extension, keep folder path
                rel_path.strip_suffix(".vb")
                    .or_else(|| rel_path.strip_suffix(".VB"))
                    .unwrap_or(&rel_path)
                    .to_string()
            } else {
                Path::new(&mod_path).file_stem().unwrap_or_default().to_string_lossy().to_string()
            };
             project.add_code_file(crate::project::CodeFile { name, code: content });
        }
    }

    // Load EmbeddedResource (.resx) files
    eprintln!("[DEBUG] Loading {} resource files...", resource_paths.len());
    for (resx_rel, dependent_upon) in &resource_paths {
        let resx_path = parent_dir.join(resx_rel);
        eprintln!("[DEBUG] Resource file: {:?} exists={}", resx_path, resx_path.exists());
        if !resx_path.exists() {
            eprintln!("[WARN] Resource file not found: {:?}", resx_path);
            continue;
        }
        match ResourceManager::load_from_file(&resx_path) {
            Ok(mut rm) => {
                // Infer a display name from the file stem
                let resx_stem = resx_path.file_stem().unwrap_or_default().to_string_lossy().to_string();
                rm.name = resx_stem.clone();
                eprintln!("[DEBUG] Loaded resource '{}': {} items, file_path={:?}", resx_rel, rm.resources.len(), rm.file_path);

                if let Some(dep) = dependent_upon {
                    // Form-dependent resource: attach to matching form
                    let dep_stem = Path::new(dep).file_stem().unwrap_or_default().to_string_lossy().to_string();
                    if let Some(fm) = project.forms.iter_mut().find(|f| f.form.name.eq_ignore_ascii_case(&dep_stem)) {
                        eprintln!("[DEBUG] Attached resource '{}' to form '{}', file_path={:?}", resx_rel, fm.form.name, rm.file_path);
                        fm.resources = rm;
                    } else {
                        eprintln!("[DEBUG] DependentUpon form '{}' not found for resource '{}', adding as project resource", dep_stem, resx_rel);
                        project.resource_files.push(rm);
                    }
                } else {
                    // No DependentUpon — project-level resource (e.g. My Project/Resources.resx)
                    eprintln!("[DEBUG] Added project resource '{}' with {} items, file_path={:?}", resx_rel, rm.resources.len(), rm.file_path);
                    project.resource_files.push(rm);
                }
            }
            Err(e) => eprintln!("[WARN] Failed to load resource {}: {}", resx_rel, e),
        }
    }

    eprintln!("[DEBUG] After resource loading: {} project resource_files, forms resources: {:?}",
        project.resource_files.len(),
        project.forms.iter().map(|f| format!("{}:file_path={:?}", f.form.name, f.resources.file_path)).collect::<Vec<_>>()
    );

    // If no explicit startup object was specified, scan code files for Sub Main
    if matches!(project.startup_object, crate::project::StartupObject::None) {
        for cf in &project.code_files {
            let upper = cf.code.to_uppercase();
            if upper.contains("SUB MAIN") {
                eprintln!("[DEBUG] Auto-detected Sub Main in code file '{}'", cf.name);
                project.startup_object = crate::project::StartupObject::SubMain;
                break;
            }
        }
    }

    // If startup_object is Form("X") but no form with that name was loaded,
    // it's actually a module reference (e.g. <StartupObject>Program</StartupObject>).
    // Check if the referenced code file contains Sub Main.
    if let crate::project::StartupObject::Form(ref name) = project.startup_object {
        let has_form = project.forms.iter().any(|f| f.form.name.eq_ignore_ascii_case(name));
        if !has_form {
            // Not a real form — check code files for Sub Main
            let has_sub_main = project.code_files.iter().any(|cf| {
                cf.code.to_uppercase().contains("SUB MAIN")
            });
            if has_sub_main {
                eprintln!("[DEBUG] '{}' is not a form, found Sub Main in code — setting SubMain", name);
                project.startup_object = crate::project::StartupObject::SubMain;
                project.startup_form = None;
            }
        }
    }

    eprintln!("[DEBUG] Project startup_object after loading: {:?}", project.startup_object);
    eprintln!("[DEBUG] Project startup_form after loading: {:?}", project.startup_form);

    // Load referenced sub-projects (<ProjectReference>)
    if !project_ref_paths.is_empty() {
        eprintln!("[DEBUG] Loading {} project references...", project_ref_paths.len());
        for ref_rel in &project_ref_paths {
            let ref_path = parent_dir.join(ref_rel);
            eprintln!("[DEBUG] Loading project reference: {:?}", ref_path);
            if !ref_path.exists() {
                eprintln!("[WARN] Referenced project not found: {:?}", ref_path);
                continue;
            }
            match load_project_vbproj(&ref_path) {
                Ok(sub_project) => {
                    eprintln!("[DEBUG] Loaded sub-project '{}': {} forms, {} code files",
                        sub_project.name, sub_project.forms.len(), sub_project.code_files.len());
                    // Track the reference name
                    project.project_references.push(sub_project.name.clone());
                    // Merge sub-project contents into main project
                    for fm in sub_project.forms {
                        if !project.forms.iter().any(|f| f.form.name == fm.form.name) {
                            project.forms.push(fm);
                        }
                    }
                    for cf in sub_project.code_files {
                        if !project.code_files.iter().any(|c| c.name == cf.name) {
                            project.code_files.push(cf);
                        }
                    }
                    for rf in sub_project.resource_files {
                        project.resource_files.push(rf);
                    }
                }
                Err(e) => {
                    eprintln!("[WARN] Failed to load referenced project {}: {}", ref_rel, e);
                }
            }
        }
    }

    Ok(project)
}

// End of serialization.rs
