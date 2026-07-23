//! Per-language WinForms designer registry — the same registration style as
//! `vybe_plugin::registry`, but for form modules (which need `Form`/`GuiState`).
//! VB / C# register themselves; the compiler/project loader reads by name.

use std::sync::{Mutex, OnceLock};

use vybe_host::GuiState;

use crate::winforms::form::Form;

#[derive(Clone, Copy)]
pub struct FormModuleLanguage {
    pub name: &'static str,
    pub load_designer: fn(&str, &mut GuiState) -> Result<(), String>,
    pub save_designer: fn(&mut GuiState, &str) -> String,
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
