//! Per-language WinForms designer registry — the same registration style as
//! `vybe_runtime::registry`, but for form modules (which need `Form`/`GuiState`).
//! VB / C# register themselves; the primitives/project loader reads by name.

use std::sync::{Mutex, OnceLock};

use crate::winforms::form::Form;

/// ⚠ CODE GENERATION ONLY. This registry answers "how does language X spell a
/// designer file", and both members take a [`Form`] — the designer's own model.
///
/// It carried two more members, `load_designer`/`save_designer`, which took a
/// widget-state object and were **registered by VB and C# but never invoked**:
/// the only consumers (`designer::project::FormModule`) call the two below.
#[derive(Clone, Copy)]
pub struct FormModuleLanguage {
    pub name: &'static str,
    pub generate_designer_code: fn(&Form) -> String,
    pub generate_user_code_stub: fn(&str) -> String,
}

fn registry() -> &'static Mutex<Vec<FormModuleLanguage>> {
    static R: OnceLock<Mutex<Vec<FormModuleLanguage>>> = OnceLock::new();
    R.get_or_init(|| Mutex::new(Vec::new()))
}

/// Register a language's WinForms designer module (idempotent by name).
pub fn register(module: FormModuleLanguage) {
    let mut r = registry().lock().unwrap();
    if !r.iter().any(|m| m.name == module.name) {
        r.push(module);
    }
}

pub fn all() -> Vec<FormModuleLanguage> {
    registry().lock().unwrap().clone()
}

pub fn find_by_name(name: &str) -> Option<FormModuleLanguage> {
    registry()
        .lock()
        .unwrap()
        .iter()
        .find(|m| m.name == name)
        .copied()
}
