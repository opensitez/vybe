use super::errors::{SaveError, SaveResult};
use super::form::Form;
use super::project::{FormModule, Project, StartupObject};
use super::resources::ResourceManager;
use super::winforms::{load_form_vb, save_form_vb};
use crate::projects::encoding::read_text_file;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::fs;
use std::path::Path;

pub fn save_project_auto(project: &Project, path: impl AsRef<Path>) -> SaveResult<()> {
    let path = path.as_ref();
    if let Some(ext) = path.extension() {
        if !ext.eq_ignore_ascii_case("vbproj") {
            return Err(SaveError::Parse(
                "Only .vbproj files are supported for saving".to_string(),
            ));
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
    xml.push_str(&format!(
        "    <RootNamespace>{}</RootNamespace>\n",
        project.name
    ));
    xml.push_str(&format!(
        "    <AssemblyName>{}</AssemblyName>\n",
        project.name
    ));
    xml.push_str("    <TargetFramework>net6.0-windows</TargetFramework>\n");
    xml.push_str("    <UseWindowsForms>true</UseWindowsForms>\n");
    match &project.startup_object {
        StartupObject::Form(form_name) => {
            xml.push_str(&format!(
                "    <StartupObject>{}.{}</StartupObject>\n",
                project.name, form_name
            ));
        }
        StartupObject::SubMain => {
            xml.push_str("    <StartupObject>Sub Main</StartupObject>\n");
        }
        StartupObject::None => {}
    }
    xml.push_str("  </PropertyGroup>\n");
    xml.push_str("  <ItemGroup>\n");
    for form_mod in &project.forms {
        let form_name = &form_mod.form.name;
        xml.push_str(&format!("    <Compile Include=\"{}.vb\">\n", form_name));
        xml.push_str("      <SubType>Form</SubType>\n");
        xml.push_str("    </Compile>\n");
        xml.push_str(&format!(
            "    <Compile Include=\"{}.Designer.vb\">\n",
            form_name
        ));
        xml.push_str(&format!(
            "      <DependentUpon>{}.vb</DependentUpon>\n",
            form_name
        ));
        xml.push_str("    </Compile>\n");
    }
    for code_file in &project.code_files {
        xml.push_str(&format!(
            "    <Compile Include=\"{}.vb\" />\n",
            code_file.name
        ));
    }
    xml.push_str("  </ItemGroup>\n");
    xml.push_str("</Project>\n");

    fs::write(path, xml)?;

    let parent_dir = path.parent().unwrap_or(Path::new("."));
    for form_mod in &project.forms {
        let mut fm = form_mod.clone();
        fm.sync_designer_code();
        save_form_vb(&fm, parent_dir)?;
    }

    for code_file in &project.code_files {
        let mod_path = parent_dir.join(format!("{}.vb", code_file.name));
        fs::write(&mod_path, &code_file.code)?;
    }

    Ok(())
}

pub fn load_project_auto(path: impl AsRef<Path>) -> SaveResult<Project> {
    let path = path.as_ref();
    if let Some(ext) = path.extension() {
        if ext.eq_ignore_ascii_case("vbproj") {
            return load_project_vbproj(path);
        }
    }
    Err(SaveError::Parse(
        "Only .vbproj files are supported".to_string(),
    ))
}

pub fn load_project_vbproj(path: impl AsRef<Path>) -> SaveResult<Project> {
    let path = path.as_ref();
    let content = read_text_file(path)?;

    let trimmed = content.trim();
    if !trimmed.starts_with('<') {
        return Err(SaveError::Parse(
            "File is not a valid XML .vbproj".to_string(),
        ));
    }

    let mut reader = Reader::from_str(&content);
    reader.config_mut().trim_text(true);

    let mut project_name = String::new();
    let mut startup_object = None;
    let mut form_paths = Vec::new();
    let mut module_paths = Vec::new();
    let mut resource_paths: Vec<(String, Option<String>)> = Vec::new();
    let mut project_ref_paths: Vec<String> = Vec::new();

    let mut current_file_path = String::new();
    let mut current_subtype = String::new();
    let mut current_dependent_upon = String::new();
    let mut capture_text = false;
    let mut current_tag = Vec::new();
    let mut in_compile = false;
    let mut in_embedded_resource = false;

    fn get_include(e: &quick_xml::events::BytesStart<'_>) -> String {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == b"Include" {
                return String::from_utf8_lossy(&attr.value).into_owned();
            }
        }
        String::new()
    }

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.name().as_ref().to_vec();
                if name == b"AssemblyName"
                    || name == b"StartupObject"
                    || name == b"SubType"
                    || name == b"DependentUpon"
                {
                    capture_text = true;
                    current_tag = name.clone();
                } else if name == b"Compile" {
                    in_compile = true;
                    current_file_path = get_include(e);
                    current_subtype = String::new();
                    current_dependent_upon = String::new();
                } else if name == b"EmbeddedResource" {
                    in_embedded_resource = true;
                    current_file_path = get_include(e);
                    current_dependent_upon = String::new();
                } else if name == b"ProjectReference" {
                    let inc = get_include(e);
                    if !inc.is_empty() {
                        project_ref_paths.push(inc.replace('\\', "/"));
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                if name.as_ref() == b"Compile" {
                    let file_path = get_include(e);
                    if !file_path.is_empty() {
                        let clean_path = file_path.replace('\\', "/");
                        if !clean_path.ends_with(".Designer.vb") {
                            module_paths.push(clean_path);
                        }
                    }
                } else if name.as_ref() == b"EmbeddedResource" {
                    let file_path = get_include(e);
                    if !file_path.is_empty() {
                        let clean_path = file_path.replace('\\', "/");
                        resource_paths.push((clean_path, None));
                    }
                } else if name.as_ref() == b"ProjectReference" {
                    let inc = get_include(e);
                    if !inc.is_empty() {
                        project_ref_paths.push(inc.replace('\\', "/"));
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if capture_text {
                    let txt = String::from_utf8_lossy(&e).into_owned();
                    if current_tag == b"AssemblyName" {
                        project_name = txt;
                    } else if current_tag == b"StartupObject" {
                        if txt == "Sub Main" || txt.is_empty() {
                            startup_object = Some("Sub Main".to_string());
                        } else if !txt.contains("My.MyApplication") {
                            startup_object = Some(txt);
                        } else {
                            startup_object = Some("__MY_APPLICATION__".to_string());
                        }
                    } else if current_tag == b"SubType" && in_compile {
                        current_subtype = txt;
                    } else if current_tag == b"DependentUpon"
                        && (in_compile || in_embedded_resource)
                    {
                        current_dependent_upon = txt;
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let qname = e.name();
                let name = qname.as_ref();
                if name == b"AssemblyName"
                    || name == b"StartupObject"
                    || name == b"SubType"
                    || name == b"DependentUpon"
                {
                    capture_text = false;
                } else if name == b"Compile" {
                    in_compile = false;
                    if !current_file_path.is_empty() {
                        let clean_path = current_file_path.replace('\\', "/");
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
        project_name = path
            .file_stem()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();
    }

    let mut project = Project::new(&project_name);

    if let Some(ref startup_str) = startup_object {
        if startup_str == "Sub Main" {
            project.startup_object = StartupObject::SubMain;
        } else {
            let form_name = if let Some(dot_pos) = startup_str.rfind('.') {
                startup_str[dot_pos + 1..].to_string()
            } else {
                startup_str.clone()
            };
            project.startup_object = StartupObject::Form(form_name.clone());
            project.startup_form = Some(form_name);
        }
    } else {
        project.startup_object = StartupObject::None;
    }

    let parent_dir = path.parent().unwrap_or(Path::new("."));

    // If startup is My.MyApplication, read MainForm from Application.myapp
    if startup_object.as_deref() == Some("__MY_APPLICATION__") {
        let myapp_path = parent_dir.join("My Project/Application.myapp");
        if myapp_path.exists() {
            if let Ok(myapp_content) = read_text_file(&myapp_path) {
                if let Some(start) = myapp_content.find("<MainForm>") {
                    let after = &myapp_content[start + 10..];
                    if let Some(end) = after.find("</MainForm>") {
                        let main_form = after[..end].trim().to_string();
                        if !main_form.is_empty() {
                            project.startup_object = StartupObject::Form(main_form.clone());
                            project.startup_form = Some(main_form);
                        }
                    }
                }
            }
        }
    }

    // SDK-style projects have no <Compile> items — auto-discover all .vb files
    if form_paths.is_empty() && module_paths.is_empty() {
        fn collect_vb_files(
            dir: &Path,
            base: &Path,
            form_paths: &mut Vec<String>,
            module_paths: &mut Vec<String>,
        ) {
            if let Ok(entries) = fs::read_dir(dir) {
                for entry in entries.flatten() {
                    let p = entry.path();
                    if p.is_dir() {
                        let dir_name = p
                            .file_name()
                            .unwrap_or_default()
                            .to_string_lossy()
                            .to_lowercase();
                        if dir_name == "bin" || dir_name == "obj" || dir_name == ".git" {
                            continue;
                        }
                        collect_vb_files(&p, base, form_paths, module_paths);
                    } else if let Some(ext) = p.extension() {
                        if ext.eq_ignore_ascii_case("vb") {
                            let rel = p
                                .strip_prefix(base)
                                .unwrap_or(&p)
                                .to_string_lossy()
                                .to_string();
                            if rel.to_lowercase().ends_with(".designer.vb") {
                                continue;
                            }
                            if let Ok(raw) = crate::projects::encoding::read_text_file(&p) {
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
    }

    // Load Forms
    for rel_path in &form_paths {
        let form_path = parent_dir.join(rel_path);
        match load_form_vb(&form_path) {
            Ok(form_module) => {
                let form_name = form_module.form.name.clone();
                project.forms.push(form_module);

                if let StartupObject::Form(ref startup_name) = project.startup_object {
                    if startup_name.ends_with(&form_name) {
                        project.startup_object = StartupObject::Form(form_name.clone());
                        project.startup_form = Some(form_name);
                    }
                }
            }
            Err(_e) => {
                let stem = Path::new(rel_path)
                    .file_stem()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .to_string();
                let user_code = read_text_file(&form_path).unwrap_or_default();
                let designer_path2 = parent_dir.join(format!("{}.Designer.vb", stem));
                let designer_code = if designer_path2.exists() {
                    read_text_file(&designer_path2).unwrap_or_default()
                } else {
                    String::new()
                };
                let form = Form::new(&stem);
                let fm = FormModule::new_vbnet(form, designer_code, user_code);
                let form_name = stem.clone();
                project.forms.push(fm);

                if let StartupObject::Form(ref startup_name) = project.startup_object {
                    if startup_name.ends_with(&form_name) {
                        project.startup_object = StartupObject::Form(form_name.clone());
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
            let name = if rel_path.contains('/') {
                rel_path
                    .strip_suffix(".vb")
                    .or_else(|| rel_path.strip_suffix(".VB"))
                    .unwrap_or(&rel_path)
                    .to_string()
            } else {
                Path::new(&mod_path)
                    .file_stem()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .to_string()
            };
            project.add_code_file(super::project::CodeFile {
                name,
                code: content,
            });
        }
    }

    // Load EmbeddedResource (.resx) files
    for (resx_rel, dependent_upon) in &resource_paths {
        let resx_path = parent_dir.join(resx_rel);
        if !resx_path.exists() {
            continue;
        }
        match ResourceManager::load_from_file(&resx_path) {
            Ok(mut rm) => {
                let resx_stem = resx_path
                    .file_stem()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .to_string();
                rm.name = resx_stem;

                if let Some(dep) = dependent_upon {
                    let dep_stem = Path::new(dep)
                        .file_stem()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string();
                    if let Some(fm) = project
                        .forms
                        .iter_mut()
                        .find(|f| f.form.name.eq_ignore_ascii_case(&dep_stem))
                    {
                        fm.resources = rm;
                    } else {
                        project.resource_files.push(rm);
                    }
                } else {
                    project.resource_files.push(rm);
                }
            }
            Err(_e) => {}
        }
    }

    // If no explicit startup object, scan code files for Sub Main
    if matches!(project.startup_object, StartupObject::None) {
        for cf in &project.code_files {
            let upper = cf.code.to_uppercase();
            if upper.contains("SUB MAIN") {
                project.startup_object = StartupObject::SubMain;
                break;
            }
        }
    }

    // If startup_object is Form("X") but no form with that name was loaded, check for Sub Main
    if let StartupObject::Form(ref name) = project.startup_object {
        let has_form = project
            .forms
            .iter()
            .any(|f| f.form.name.eq_ignore_ascii_case(name));
        if !has_form {
            let has_sub_main = project
                .code_files
                .iter()
                .any(|cf| cf.code.to_uppercase().contains("SUB MAIN"));
            if has_sub_main {
                project.startup_object = StartupObject::SubMain;
                project.startup_form = None;
            }
        }
    }

    // Load referenced sub-projects (<ProjectReference>)
    if !project_ref_paths.is_empty() {
        for ref_rel in &project_ref_paths {
            let ref_path = parent_dir.join(ref_rel);
            if !ref_path.exists() {
                continue;
            }
            match load_project_vbproj(&ref_path) {
                Ok(sub_project) => {
                    project.project_references.push(sub_project.name.clone());
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
                Err(_e) => {}
            }
        }
    }

    Ok(project)
}
