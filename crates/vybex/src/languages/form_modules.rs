use crate::projects::vbforms::form::Form;
use vybe_host::GuiState;

/// Language-registered form module surface.
///
/// Mirrors `languages::Language` registration style so designer/emission
/// support can scale per-language without hard-coding VB paths.
pub struct FormModuleLanguage {
    pub name: &'static str,
    pub load_designer: fn(&str, &mut GuiState) -> Result<(), String>,
    pub save_designer: fn(&mut GuiState, &str) -> String,
    pub generate_designer_code: fn(&Form) -> String,
    pub generate_user_code_stub: fn(&str) -> String,
}

/// All registered form modules.
pub fn all() -> Vec<FormModuleLanguage> {
    vec![
        FormModuleLanguage {
            name: "vb",
            load_designer: crate::languages::vb::forms::load_designer,
            save_designer: crate::languages::vb::forms::save_designer,
            generate_designer_code: crate::languages::vb::designer_codegen::generate_designer_code,
            generate_user_code_stub: crate::languages::vb::designer_codegen::generate_user_code_stub,
        },
        FormModuleLanguage {
            name: "csharp",
            load_designer: crate::languages::csharp::forms::load_designer,
            save_designer: crate::languages::csharp::forms::save_designer,
            generate_designer_code: crate::languages::csharp::forms::generate_designer_code,
            generate_user_code_stub: crate::languages::csharp::forms::generate_user_code_stub,
        },
    ]
}

/// Find a registered form module by language name.
pub fn find_by_name(name: &str) -> Option<FormModuleLanguage> {
    all().into_iter().find(|m| m.name == name)
}
