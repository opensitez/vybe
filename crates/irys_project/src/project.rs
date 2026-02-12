use irys_forms::Form;
use serde::{Deserialize, Serialize};
use crate::resources::ResourceManager;

/// Specifies what the project starts with
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum StartupObject {
    /// Start with a form
    Form(String),
    /// Start with Sub Main in a module
    SubMain,
    /// No startup object specified
    None,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CodeFile {
    pub name: String,
    pub code: String,
}

impl CodeFile {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            code: String::new(),
        }
    }
}

/// Backward compatibility aliases
pub type VBModule = CodeFile;
pub type ClassModule = CodeFile;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum FormFormat {
    Classic,
    VbNet {
        designer_code: String,
        user_code: String,
    },
}

impl Default for FormFormat {
    fn default() -> Self {
        FormFormat::VbNet {
            designer_code: String::new(),
            user_code: String::new(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FormModule {
    pub form: Form,
    pub code: String,
    #[serde(default)]
    pub format: FormFormat,
    /// Form-specific resources (Form1.resx) — localizable strings, embedded icons, etc.
    #[serde(default)]
    pub resources: ResourceManager,
}

impl FormModule {
    pub fn new(form: Form) -> Self {
        let name = form.name.clone();
        let designer_code = irys_forms::serialization::designer_codegen::generate_designer_code(&form);
        let user_code = irys_forms::serialization::designer_codegen::generate_user_code_stub(&name);
        Self {
            form,
            code: String::new(),
            format: FormFormat::VbNet {
                designer_code,
                user_code,
            },
            resources: ResourceManager::new_named(format!("{name}")),
        }
    }

    /// Create a classic (VB6-style) form module
    pub fn new_classic(form: Form) -> Self {
        let name = form.name.clone();
        Self {
            form,
            code: String::new(),
            format: FormFormat::Classic,
            resources: ResourceManager::new_named(format!("{name}")),
        }
    }

    pub fn new_vbnet(form: Form, designer_code: String, user_code: String) -> Self {
        let name = form.name.clone();
        Self {
            form,
            code: String::new(),
            format: FormFormat::VbNet {
                designer_code,
                user_code,
            },
            resources: ResourceManager::new_named(format!("{name}")),
        }
    }

    /// Returns the user-editable code (Classic: code field, VbNet: user_code).
    pub fn get_user_code(&self) -> &str {
        match &self.format {
            FormFormat::Classic => &self.code,
            FormFormat::VbNet { user_code, .. } => user_code,
        }
    }

    /// Returns the designer-generated code (Classic: empty, VbNet: designer_code).
    pub fn get_designer_code(&self) -> &str {
        match &self.format {
            FormFormat::Classic => "",
            FormFormat::VbNet { designer_code, .. } => designer_code,
        }
    }

    /// Sets the user-editable code.
    pub fn set_user_code(&mut self, code: String) {
        match &mut self.format {
            FormFormat::Classic => self.code = code,
            FormFormat::VbNet { user_code, .. } => *user_code = code,
        }
    }

    /// Regenerates designer_code from the Form object (VbNet only).
    pub fn sync_designer_code(&mut self) {
        if let FormFormat::VbNet { designer_code, .. } = &mut self.format {
            *designer_code = irys_forms::serialization::designer_codegen::generate_designer_code(&self.form);
        }
    }

    /// Returns true if this is a VB.NET format form.
    pub fn is_vbnet(&self) -> bool {
        matches!(self.format, FormFormat::VbNet { .. })
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct Project {
    pub name: String,
    #[serde(default)]
    pub startup_object: StartupObject,
    /// Deprecated: use startup_object instead. Kept for backward compatibility.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub startup_form: Option<String>,
    pub forms: Vec<FormModule>,
    pub code_files: Vec<CodeFile>,
    /// Multiple resource files (.resx) — e.g. Resources.resx, Strings.resx, Images.resx
    #[serde(default)]
    pub resource_files: Vec<ResourceManager>,
    /// Names of referenced sub-projects (from <ProjectReference> in .vbproj)
    #[serde(default)]
    pub project_references: Vec<String>,
    /// Deprecated: single resource manager. Kept for backward compat deserialization.
    #[serde(skip_serializing)]
    pub resources: ResourceManager,
}

impl Default for StartupObject {
    fn default() -> Self {
        StartupObject::None
    }
}

impl Project {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            startup_object: StartupObject::None,
            startup_form: None,
            forms: Vec::new(),
            code_files: Vec::new(),
            resource_files: Vec::new(),
            project_references: Vec::new(),
            resources: ResourceManager::new(),
        }
    }

    pub fn add_form(&mut self, form: Form) {
        if matches!(self.startup_object, StartupObject::None) {
            self.startup_object = StartupObject::Form(form.name.clone());
            self.startup_form = Some(form.name.clone());
        }
        self.forms.push(FormModule::new(form));
    }

    pub fn add_code_file(&mut self, code_file: CodeFile) {
        self.code_files.push(code_file);
    }

    pub fn remove_form(&mut self, name: &str) -> bool {
        let len = self.forms.len();
        self.forms.retain(|f| f.form.name != name);
        self.forms.len() < len
    }

    pub fn remove_code_file(&mut self, name: &str) -> bool {
        let len = self.code_files.len();
        self.code_files.retain(|cf| cf.name != name);
        self.code_files.len() < len
    }

    pub fn get_form(&self, name: &str) -> Option<&FormModule> {
        self.forms.iter().find(|f| f.form.name == name)
    }

    pub fn get_form_mut(&mut self, name: &str) -> Option<&mut FormModule> {
        self.forms.iter_mut().find(|f| f.form.name == name)
    }

    pub fn get_code_file(&self, name: &str) -> Option<&CodeFile> {
        self.code_files.iter().find(|cf| cf.name == name)
    }

    pub fn get_code_file_mut(&mut self, name: &str) -> Option<&mut CodeFile> {
        self.code_files.iter_mut().find(|cf| cf.name == name)
    }

    pub fn get_startup_form(&self) -> Option<&FormModule> {
        // Try new startup_object first, then fall back to deprecated startup_form
        match &self.startup_object {
            StartupObject::Form(name) => self.get_form(name),
            _ => self.startup_form.as_ref().and_then(|name| self.get_form(name)),
        }
    }

    /// Returns the name of the startup form, if the startup object is a form
    pub fn get_startup_form_name(&self) -> Option<&str> {
        match &self.startup_object {
            StartupObject::Form(name) => Some(name.as_str()),
            _ => self.startup_form.as_deref(),
        }
    }

    /// Returns true if the project starts with Sub Main
    pub fn starts_with_main(&self) -> bool {
        matches!(self.startup_object, StartupObject::SubMain)
    }
}

// Custom Deserialize for backward compatibility with old JSON projects
// that have separate "modules" and "classes" fields
impl<'de> Deserialize<'de> for Project {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        #[derive(Deserialize)]
        struct ProjectHelper {
            name: String,
            #[serde(default)]
            startup_object: Option<StartupObject>,
            #[serde(default)]
            startup_form: Option<String>,
            #[serde(default)]
            forms: Vec<FormModule>,
            #[serde(default)]
            code_files: Vec<CodeFile>,
            #[serde(default)]
            modules: Vec<CodeFile>,
            #[serde(default)]
            classes: Vec<CodeFile>,
            #[serde(default)]
            resource_files: Vec<ResourceManager>,
            #[serde(default)]
            resources: ResourceManager,
        }

        let helper = ProjectHelper::deserialize(deserializer)?;

        let mut code_files = helper.code_files;
        if code_files.is_empty() && (!helper.modules.is_empty() || !helper.classes.is_empty()) {
            code_files.extend(helper.modules);
            code_files.extend(helper.classes);
        }

        // Migrate old startup_form to new startup_object
        let startup_object = if let Some(so) = helper.startup_object {
            // New format: use the startup_object directly
            so
        } else if let Some(ref form_name) = helper.startup_form {
            // Old format: convert startup_form to startup_object
            StartupObject::Form(form_name.clone())
        } else {
            StartupObject::None
        };

        // Migrate old single resources to resource_files if needed
        let resource_files = if !helper.resource_files.is_empty() {
            helper.resource_files
        } else if !helper.resources.resources.is_empty() {
            vec![helper.resources.clone()]
        } else {
            Vec::new()
        };

        Ok(Project {
            name: helper.name,
            startup_object,
            startup_form: helper.startup_form,
            forms: helper.forms,
            code_files,
            resource_files,
            project_references: Vec::new(),
            resources: helper.resources,
        })
    }
}
