use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::fs;
use std::path::Path;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LanguageDef {
    #[serde(default)]
    pub keywords: HashSet<String>,
    #[serde(rename = "typeKeywords", default)]
    pub type_keywords: HashSet<String>,
    #[serde(default)]
    pub constants: HashSet<String>,
    #[serde(default)]
    pub operators: HashSet<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct LanguageJson {
    pub language: LanguageDef,
}

pub fn load_language(name: &str) -> Option<LanguageDef> {
    // Current directory is vybe/crates/code_editor
    let path = Path::new("basic-languages").join(name).join(format!("{}.json", name));
    if let Ok(content) = fs::read_to_string(path) {
        if let Ok(json) = serde_json::from_str::<LanguageJson>(&content) {
            return Some(json.language);
        }
    }
    None
}
