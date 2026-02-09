use irys_forms::Form;
use serde::{Deserialize, Serialize};
use crate::resources::ResourceManager;

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
        FormFormat::Classic
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FormModule {
    pub form: Form,
    pub code: String,
    #[serde(default)]
    pub format: FormFormat,
}

impl FormModule {
    pub fn new(form: Form) -> Self {
        Self {
            form,
            code: String::new(),
            format: FormFormat::Classic,
        }
    }

    pub fn new_vbnet(form: Form, designer_code: String, user_code: String) -> Self {
        Self {
            form,
            code: String::new(),
            format: FormFormat::VbNet {
                designer_code,
                user_code,
            },
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
    pub startup_form: Option<String>,
    pub forms: Vec<FormModule>,
    pub code_files: Vec<CodeFile>,
    #[serde(default)]
    pub resources: ResourceManager,
}

impl Project {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            startup_form: None,
            forms: Vec::new(),
            code_files: Vec::new(),
            resources: ResourceManager::new(),
        }
    }

    pub fn add_form(&mut self, form: Form) {
        if self.startup_form.is_none() {
            self.startup_form = Some(form.name.clone());
        }
        self.forms.push(FormModule::new(form));
    }

    pub fn add_code_file(&mut self, code_file: CodeFile) {
        self.code_files.push(code_file);
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
        self.startup_form.as_ref().and_then(|name| self.get_form(name))
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
            resources: ResourceManager,
        }

        let helper = ProjectHelper::deserialize(deserializer)?;

        let mut code_files = helper.code_files;
        if code_files.is_empty() && (!helper.modules.is_empty() || !helper.classes.is_empty()) {
            code_files.extend(helper.modules);
            code_files.extend(helper.classes);
        }

        Ok(Project {
            name: helper.name,
            startup_form: helper.startup_form,
            forms: helper.forms,
            code_files,
            resources: helper.resources,
        })
    }
}
