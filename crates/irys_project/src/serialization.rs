
use crate::project::{Project, FormModule, FormFormat, CodeFile};
use crate::classicforms::load_form_frm;
use crate::winforms::{load_form_vb, save_form_vb};
use crate::errors::{SaveError, SaveResult};
use crate::resources::ResourceManager;
use irys_forms::{Control, ControlType, Form};
use std::path::Path;
use std::fs;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use crate::encoding::read_text_file;

pub fn save_project_auto(project: &Project, path: impl AsRef<Path>) -> SaveResult<()> {
    let path = path.as_ref();
    if let Some(ext) = path.extension() {
        if ext.eq_ignore_ascii_case("vbp") {
            return save_project_vbp(project, path);
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
        ControlType::BindingNavigator => "VB.ToolBar",
        ControlType::TabControl => "VB.TabControl",
        ControlType::TabPage => "VB.TabPage",
        ControlType::ProgressBar => "VB.ProgressBar",
        ControlType::NumericUpDown => "VB.NumericUpDown",
        ControlType::MenuStrip => "VB.MenuStrip",
        ControlType::ToolStripMenuItem => "VB.ToolStripMenuItem",
        ControlType::ContextMenuStrip => "VB.ContextMenuStrip",
        ControlType::StatusStrip => "VB.StatusStrip",
        ControlType::ToolStripStatusLabel => "VB.ToolStripStatusLabel",
        ControlType::DateTimePicker => "VB.DateTimePicker",
        ControlType::LinkLabel => "VB.LinkLabel",
        ControlType::ToolStrip => "VB.ToolStrip",
        ControlType::TrackBar => "VB.TrackBar",
        ControlType::MaskedTextBox => "VB.MaskedTextBox",
        ControlType::SplitContainer => "VB.SplitContainer",
        ControlType::FlowLayoutPanel => "VB.FlowLayoutPanel",
        ControlType::TableLayoutPanel => "VB.TableLayoutPanel",
        ControlType::MonthCalendar => "VB.MonthCalendar",
        ControlType::HScrollBar => "VB.HScrollBar",
        ControlType::VScrollBar => "VB.VScrollBar",
        ControlType::ToolTip => "VB.ToolTip",
        ControlType::BindingSourceComponent => "VB.BindingSource",
        ControlType::DataSetComponent => "VB.DataSet",
        ControlType::DataTableComponent => "VB.DataTable",
        ControlType::DataAdapterComponent => "VB.DataAdapter",
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
        ControlType::DataAdapterComponent => {
            if let Some(cs) = control.properties.get_string("ConnectionString") {
                content.push_str(&format!("{}ConnectionString =   \"{}\"\n", inner_indent, escape_quotes(cs)));
            }
            if let Some(sc) = control.properties.get_string("SelectCommand") {
                content.push_str(&format!("{}SelectCommand   =   \"{}\"\n", inner_indent, escape_quotes(sc)));
            }
        }
        ControlType::BindingSourceComponent => {
            if let Some(ds) = control.properties.get_string("DataSource") {
                content.push_str(&format!("{}DataSource      =   \"{}\"\n", inner_indent, escape_quotes(ds)));
            }
            if let Some(dm) = control.properties.get_string("DataMember") {
                content.push_str(&format!("{}DataMember      =   \"{}\"\n", inner_indent, escape_quotes(dm)));
            }
        }
        ControlType::DataSetComponent => {
            if let Some(dsn) = control.properties.get_string("DataSetName") {
                content.push_str(&format!("{}DataSetName     =   \"{}\"\n", inner_indent, escape_quotes(dsn)));
            }
        }
        ControlType::DataTableComponent => {
            if let Some(tn) = control.properties.get_string("TableName") {
                content.push_str(&format!("{}TableName       =   \"{}\"\n", inner_indent, escape_quotes(tn)));
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
    let content = read_text_file(path)?;
    
    // Detect if this is actually a VBP-style file disguised as .vbproj
    // Real XML starts with '<' (possibly after whitespace/BOM)
    let trimmed = content.trim();
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

pub fn load_project_vbp(path: impl AsRef<Path>) -> SaveResult<Project> {
    let path = path.as_ref();
    let content = read_text_file(path)?;
    
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
    let content = read_text_file(path)?;
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
