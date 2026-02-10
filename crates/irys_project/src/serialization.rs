
use crate::project::{Project, FormModule, FormFormat, CodeFile};
use crate::classicforms::load_form_frm;
use crate::winforms::{load_form_vb, save_form_vb};
use crate::errors::{SaveError, SaveResult};
use irys_forms::{Control, ControlType};
use std::path::Path;
use std::fs;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

pub fn save_project_auto(project: &Project, path: impl AsRef<Path>) -> SaveResult<()> {
    let path = path.as_ref();
    if let Some(ext) = path.extension() {
        if ext.eq_ignore_ascii_case("vbproj") {
            return save_project_vbproj(project, path);
        }
    }
    save_project_vbp(project, path)
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

pub fn save_project_vbp(project: &Project, path: impl AsRef<Path>) -> SaveResult<()> {
    let path = path.as_ref();
    let mut vbp_content = String::new();

    // Basic Header
    vbp_content.push_str("Type=Exe\n");
    // Standard Reference (OLE Automation)
    vbp_content.push_str("Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\\..\\..\\WINDOWS\\SYSTEM\\stdole2.tlb#OLE Automation\n");

    // Forms
    for form in &project.forms {
        match &form.format {
            FormFormat::Classic => {
                vbp_content.push_str(&format!("Form={}.frm\n", form.form.name));
            }
            FormFormat::VbNet { .. } => {
                vbp_content.push_str(&format!("FormVB={}.vb\n", form.form.name));
            }
        }
    }

    // Code files
    for code_file in &project.code_files {
         vbp_content.push_str(&format!("Module={}; {}.bas\n", code_file.name, code_file.name));
    }

    // Project Properties
    match &project.startup_object {
        crate::project::StartupObject::Form(form_name) => {
            vbp_content.push_str(&format!("Startup=\"{}\"\n", form_name));
        }
        crate::project::StartupObject::SubMain => {
            vbp_content.push_str("Startup=\"Sub Main\"\n");
        }
        crate::project::StartupObject::None => {
            vbp_content.push_str("Startup=\"Sub Main\"\n");
        }
    }
    
    vbp_content.push_str("Command32=\"\"\n");
    vbp_content.push_str(&format!("Name=\"{}\"\n", project.name));
    vbp_content.push_str("HelpContextID=\"0\"\n");
    vbp_content.push_str("CompatibleMode=\"0\"\n");
    vbp_content.push_str("MajorVer=1\n");
    vbp_content.push_str("MinorVer=0\n");
    vbp_content.push_str("RevisionVer=0\n");
    vbp_content.push_str("AutoIncrementVer=0\n");
    vbp_content.push_str("ServerSupportFiles=0\n");
    vbp_content.push_str("CompilationType=0\n");
    vbp_content.push_str("OptimizationType=0\n");
    vbp_content.push_str("FavorPentiumPro(tm)=0\n");
    vbp_content.push_str("CodeViewDebugInfo=0\n");
    vbp_content.push_str("NoAliasing=0\n");
    vbp_content.push_str("BoundsCheck=0\n");
    vbp_content.push_str("OverflowCheck=0\n");
    vbp_content.push_str("FlPointCheck=0\n");
    vbp_content.push_str("FDIVCheck=0\n");
    vbp_content.push_str("UnroundedFP=0\n");
    vbp_content.push_str("StartMode=0\n");
    vbp_content.push_str("Unattended=0\n");
    vbp_content.push_str("Retained=0\n");
    vbp_content.push_str("ThreadPerObject=0\n");
    vbp_content.push_str("MaxNumberOfThreads=1\n");

    fs::write(path, vbp_content)?;

    // Save Forms
    let parent_dir = path.parent().unwrap_or(Path::new("."));
    for form_mod in &project.forms {
        match &form_mod.format {
            FormFormat::Classic => {
                let form_path = parent_dir.join(format!("{}.frm", form_mod.form.name));
                save_form_from_module(form_mod, &form_path)?;
            }
            FormFormat::VbNet { .. } => {
                let mut fm = form_mod.clone();
                fm.sync_designer_code();
                save_form_vb(&fm, parent_dir)?;
            }
        }
    }

    // Save code files
    for code_file in &project.code_files {
        let mod_path = parent_dir.join(format!("{}.bas", code_file.name));
        save_code_file(code_file, &mod_path)?;
    }

    Ok(())
}

fn save_form_from_module(module: &FormModule, path: &Path) -> SaveResult<()> {
    let mut content = String::new();
    let form = &module.form;

    // Header
    content.push_str("VERSION 5.00\n");
    
    // Form Container Start
    content.push_str(&format!("Begin VB.Form {}\n", form.name));
    
    // Form Properties
    content.push_str(&format!("   Caption         =   \"{}\"\n", form.caption));
    content.push_str(&format!("   ClientHeight    =   {}\n", form.height));
    content.push_str(&format!("   ClientLeft      =   60\n"));
    content.push_str(&format!("   ClientTop       =   345\n"));
    content.push_str(&format!("   ClientWidth     =   {}\n", form.width));
    content.push_str(&format!("   LinkTopic       =   \"{}\"\n", form.name));
    content.push_str(&format!("   ScaleHeight     =   {}\n", form.height));
    content.push_str(&format!("   ScaleWidth      =   {}\n", form.width));
    if let Some(bc) = form.back_color.as_ref() {
        content.push_str(&format!("   BackColor       =   {}\n", color_to_vb6(bc)));
    }
    if let Some(fc) = form.fore_color.as_ref() {
        content.push_str(&format!("   ForeColor       =   {}\n", color_to_vb6(fc)));
    }
    if let Some(font) = form.font.as_ref() {
        content.push_str(&format!("   Font            =   \"{}\"\n", font));
    }
    content.push_str("   StartUpPosition =   3  'Windows Default\n");

    // Controls
    // TODO: Handle hierarchy if we have containers like Frames
    // For now, write all controls as children of the form
    for control in &form.controls {
        write_control(&mut content, control, 1)?;
    }

    // Form Container End
    content.push_str("End\n");

    // Attributes
    content.push_str(&format!("Attribute VB_Name = \"{}\"\n", form.name));
    content.push_str("Attribute VB_GlobalNameSpace = False\n");
    content.push_str("Attribute VB_Creatable = False\n");
    content.push_str("Attribute VB_PredeclaredId = True\n");
    content.push_str("Attribute VB_Exposed = False\n");

    // Code
    content.push_str(&module.code);

    fs::write(path, content)?;
    Ok(())
}

fn save_code_file(code_file: &CodeFile, path: &Path) -> SaveResult<()> {
    let mut content = String::new();
    content.push_str(&format!("Attribute VB_Name = \"{}\"\n", code_file.name));
    content.push_str(&code_file.code);
    fs::write(path, content)?;
    Ok(())
}

fn write_control(content: &mut String, control: &Control, indent_level: usize) -> SaveResult<()> {
    let indent = "   ".repeat(indent_level);
    let vb_type = match control.control_type {
        ControlType::Button => "VB.CommandButton",
        ControlType::Label => "VB.Label",
        ControlType::TextBox => "VB.TextBox",
        ControlType::CheckBox => "VB.CheckBox",
        ControlType::RadioButton => "VB.OptionButton",
        ControlType::ComboBox => "VB.ComboBox",
        ControlType::ListBox => "VB.ListBox",
        ControlType::Frame => "VB.Frame",
        ControlType::PictureBox => "VB.PictureBox",
        ControlType::RichTextBox => "RichTextLib.RichTextBox",
        ControlType::WebBrowser => "SHDocVw.WebBrowser",
        ControlType::TreeView => "MSComctlLib.TreeView",
        ControlType::DataGridView => "MSDataGridLib.DataGrid",
        ControlType::Panel => "VB.PictureBox",
        ControlType::ListView => "MSComctlLib.ListView",
    };

    content.push_str(&format!("{}Begin {} {} \n", indent, vb_type, control.name));

    // Control Properties
    let inner_indent = "   ".repeat(indent_level + 1);

    // Index property (for control arrays)
    if let Some(idx) = control.index {
        content.push_str(&format!("{}Index           =   {}\n", inner_indent, idx));
    }

    // Common properties
    if let Some(caption) = control.get_caption() {
        content.push_str(&format!("{}Caption         =   \"{}\"\n", inner_indent, caption));
    }
    
    if let Some(text) = control.get_text() {
        content.push_str(&format!("{}Text            =   \"{}\"\n", inner_indent, text));
    }

    content.push_str(&format!("{}Height          =   {}\n", inner_indent, control.bounds.height));
    content.push_str(&format!("{}Left            =   {}\n", inner_indent, control.bounds.x));
    content.push_str(&format!("{}TabIndex        =   {}\n", inner_indent, control.tab_index));
    content.push_str(&format!("{}Top             =   {}\n", inner_indent, control.bounds.y));
    content.push_str(&format!("{}Width           =   {}\n", inner_indent, control.bounds.width));

    if let Some(bc) = control.get_back_color() {
        content.push_str(&format!("{}BackColor       =   {}\n", inner_indent, color_to_vb6(bc)));
    }
    if let Some(fc) = control.get_fore_color() {
        content.push_str(&format!("{}ForeColor       =   {}\n", inner_indent, color_to_vb6(fc)));
    }
    if let Some(font) = control.get_font() {
        content.push_str(&format!("{}Font            =   \"{}\"\n", inner_indent, font));
    }
    content.push_str(&format!("{}Enabled         =   {}\n", inner_indent, bool_to_vb6(control.is_enabled())));
    content.push_str(&format!("{}Visible         =   {}\n", inner_indent, bool_to_vb6(control.is_visible())));

    match control.control_type {
        ControlType::CheckBox | ControlType::RadioButton => {
             // In VB, 1 is Checked, 0 is Unchecked. RadioButton uses True/False sometimes?
             // Actually CheckBox: 0=Unchecked, 1=Checked, 2=Grayed
             // OptionButton: True/False
             if let Some(val) = control.properties.get_bool("Value") {
                 let v = if val { 1 } else { 0 };
                 content.push_str(&format!("{}Value           =   {}\n", inner_indent, v));
             }
        }
        ControlType::ListBox | ControlType::ComboBox => {
            let items = control.get_list_items();
            for (idx, item) in items.iter().enumerate() {
                content.push_str(&format!("{}List({})        =   \"{}\"\n", inner_indent, idx, escape_quotes(item)));
            }
            if let Some(val) = control.properties.get_int("ListIndex") {
                content.push_str(&format!("{}ListIndex       =   {}\n", inner_indent, val));
            }
            if let Some(val) = control.properties.get_string("Value") {
                content.push_str(&format!("{}Value           =   \"{}\"\n", inner_indent, escape_quotes(val)));
            }
        }
        ControlType::WebBrowser => {
            if let Some(url) = control.properties.get_string("URL") {
                content.push_str(&format!("{}URL             =   \"{}\"\n", inner_indent, escape_quotes(url)));
            }
        }
        ControlType::RichTextBox => {
            if let Some(html) = control.properties.get_string("HTML") {
                content.push_str(&format!("{}HTML            =   \"{}\"\n", inner_indent, escape_quotes(html)));
            }
            if let Some(tb) = control.properties.get_bool("ToolbarVisible") {
                content.push_str(&format!("{}ToolbarVisible  =   {}\n", inner_indent, bool_to_vb6(tb)));
            }
        }
        ControlType::TreeView => {
            if let Some(sep) = control.properties.get_string("PathSeparator") {
                content.push_str(&format!("{}PathSeparator   =   \"{}\"\n", inner_indent, escape_quotes(sep)));
            }
        }
        _ => {}
    }

    content.push_str(&format!("{}End\n", indent));
    Ok(())
}

fn color_to_vb6(color: &str) -> String {
    let c = color.trim();
    if let Some(hex) = c.strip_prefix('#') {
        if hex.len() == 6 {
            if let Ok(val) = u32::from_str_radix(hex, 16) {
                let r = (val >> 16) & 0xFF;
                let g = (val >> 8) & 0xFF;
                let b = val & 0xFF;
                return format!("&H00{:02X}{:02X}{:02X}&", b, g, r);
            }
        }
    }
    // Fallback to default light gray
    "&H00FCFAF8&".to_string()
}

fn bool_to_vb6(value: bool) -> &'static str {
    if value { "-1" } else { "0" }
}

fn escape_quotes(text: &str) -> String {
    text.replace('"', "\"\"")
}
// --- Loading ---

pub fn load_project_auto(path: impl AsRef<Path>) -> SaveResult<Project> {
    let path = path.as_ref();
    eprintln!("[DEBUG] load_project_auto: {:?}", path);
    if let Some(ext) = path.extension() {
        eprintln!("[DEBUG] extension: {:?}", ext);
        if ext.eq_ignore_ascii_case("vbproj") {
            return load_project_vbproj(path);
        }
    }
    load_project_vbp(path)
}

pub fn load_project_vbproj(path: impl AsRef<Path>) -> SaveResult<Project> {
    let path = path.as_ref();
    let content = fs::read_to_string(path)?;
    
    // Detect if this is actually a VBP-style file disguised as .vbproj
    // Real XML starts with '<' (possibly after whitespace/BOM)
    let trimmed = content.trim_start_matches('\u{feff}').trim();
    if !trimmed.starts_with('<') {
        eprintln!("[DEBUG] load_project_vbproj: not XML, falling back to VBP parser");
        return load_project_vbp(path);
    }
    
    let mut reader = Reader::from_str(&content);
    reader.trim_text(true);

    let mut project_name = String::new();
    let mut startup_object = None;
    let mut form_paths = Vec::new();
    let mut module_paths = Vec::new();

    // State tracking
    let mut current_file_path = String::new();
    let mut current_subtype = String::new();
    let mut capture_text = false;
    let mut current_tag = Vec::new();
    let mut in_compile = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = e.name().as_ref().to_vec();
                if name == b"AssemblyName" || name == b"StartupObject" || name == b"SubType" {
                    capture_text = true;
                    current_tag = name.clone();
                } else if name == b"Compile" {
                    in_compile = true;
                    current_file_path = String::new();
                    current_subtype = String::new();
                    for attr in e.attributes() {
                        if let Ok(attr) = attr {
                            if attr.key.as_ref() == b"Include" {
                                if let Ok(val) = attr.unescape_value() {
                                    current_file_path = val.into_owned();
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
                            }
                        } else if current_tag == b"SubType" && in_compile {
                            current_subtype = txt;
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let qname = e.name();
                let name = qname.as_ref();
                if name == b"AssemblyName" || name == b"StartupObject" || name == b"SubType" {
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

    eprintln!("[DEBUG] load_project_vbproj: project='{}', form_paths={:?}, module_paths={:?}", project_name, form_paths, module_paths);

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
            Err(e) => eprintln!("Failed to load form {}: {}", rel_path, e),
        }
    }

    // Load Modules
    for rel_path in module_paths {
        let mod_path = parent_dir.join(rel_path);
        // Modules in .vbproj are just .vb files, unlike .bas
        // We can reuse load_code_file_bas or create a new load_code_file_vb
        // .vb files don't have "Attribute VB_Name", they have "Module ModuleName" inside
        // For now, read content and infer name from filename
        if let Ok(content) = fs::read_to_string(&mod_path) {
            let name = Path::new(&mod_path).file_stem().unwrap_or_default().to_string_lossy().to_string();
             // TODO: parse "Module X" to get real name
             project.add_code_file(crate::project::CodeFile { name, code: content });
        }
    }

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

    eprintln!("[DEBUG] Project startup_object after loading: {:?}", project.startup_object);
    eprintln!("[DEBUG] Project startup_form after loading: {:?}", project.startup_form);

    Ok(project)
}

pub fn load_project_vbp(path: impl AsRef<Path>) -> SaveResult<Project> {
    let path = path.as_ref();
    let content = fs::read_to_string(path)?;
    
    let mut name = String::new();
    let mut startup_value = None;
    let mut form_paths = Vec::new();
    let mut vb_form_paths = Vec::new();
    let mut module_paths = Vec::new();

    for line in content.lines() {
        let line = line.trim();
        if line.starts_with("Name=\"") {
            name = line.trim_start_matches("Name=\"").trim_end_matches('"').to_string();
        } else if line.starts_with("Startup=\"") {
            let s = line.trim_start_matches("Startup=\"").trim_end_matches('"');
            startup_value = Some(s.to_string());
        } else if line.starts_with("FormVB=") {
            vb_form_paths.push(line.trim_start_matches("FormVB=").to_string());
        } else if line.starts_with("Form=") {
            let relative_path = line.trim_start_matches("Form=");
            form_paths.push(relative_path.to_string());
        } else if line.starts_with("Module=") {
             let parts: Vec<&str> = line.trim_start_matches("Module=").split(';').collect();
             if parts.len() == 2 {
                 let mod_name = parts[0].trim();
                 let mod_path = parts[1].trim();
                 module_paths.push((mod_name.to_string(), mod_path.to_string()));
             }
        }
    }

    if name.is_empty() {
        // Fallback if Name not found (e.g. minimal file)
        name = path.file_stem().unwrap_or_default().to_string_lossy().to_string();
    }

    let mut project = Project::new(&name);
    
    // Set startup object based on parsed Startup value
    if let Some(ref startup) = startup_value {
        if startup.eq_ignore_ascii_case("Sub Main") || startup == "Sub Main" {
            project.startup_object = crate::project::StartupObject::SubMain;
        } else {
            project.startup_object = crate::project::StartupObject::Form(startup.clone());
            project.startup_form = Some(startup.clone());
        }
    }

    let parent_dir = path.parent().unwrap_or(Path::new("."));

    // Load Forms
    for rel_path in form_paths {
        let form_path = parent_dir.join(rel_path);
        let form_module = load_form_frm(&form_path)?;
        project.add_form(form_module.form);
        // We need to set code separately or modify add_form to take FormModule?
        // Project.add_form takes Form. But we need to store code too.
        // Project struct has `forms: Vec<FormModule>`.
        // Wait, `project.add_form(form)` creates a default wrapper.
        // We should fix Project API or access forms directly.
        // Let's modify the last added form's code.
        if let Some(last_mod) = project.forms.last_mut() {
            last_mod.code = form_module.code;
        }
    }

    // Load VB.NET Forms
    for rel_path in vb_form_paths {
        let form_path = parent_dir.join(&rel_path);
        match load_form_vb(&form_path) {
            Ok(form_module) => {
                project.forms.push(form_module);
                if project.startup_form.is_none() {
                    if let Some(last) = project.forms.last() {
                        project.startup_form = Some(last.form.name.clone());
                    }
                }
            }
            Err(e) => {
                eprintln!("Failed to load VB.NET form {}: {}", rel_path, e);
            }
        }
    }

    // Load Modules
    for (_mod_name, rel_path) in module_paths {
        let mod_path = parent_dir.join(rel_path);
        let code_file = load_code_file_bas(&mod_path)?;
        project.add_code_file(code_file);
    }

    Ok(project)
}

fn load_code_file_bas(path: &Path) -> SaveResult<CodeFile> {
    let content = fs::read_to_string(path)?;
    let mut name = String::new();
    let mut code = String::new();

    for line in content.lines() {
        if line.trim().starts_with("Attribute VB_Name =") {
             name = line.split('=').nth(1).unwrap_or("").trim().trim_matches('"').to_string();
        } else if !line.trim().starts_with("Attribute") {
             code.push_str(line);
             code.push('\n');
        }
    }
    
    // If no VB_Name attribute found, use filename
    if name.is_empty() {
        name = path.file_stem()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();
    }

    Ok(CodeFile {
        name,
        code,
    })
}
