use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::fs;
use std::path::Path;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentDef {
    #[serde(rename = "lineComment")]
    pub line_comment: Option<String>,
    #[serde(rename = "blockComment")]
    pub block_comment: Option<(String, String)>,
}

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
    #[serde(default)]
    pub comments: Option<CommentDef>,
    #[serde(default)]
    pub brackets: Vec<(String, String)>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct LanguageJson {
    pub conf: Option<ConfJson>,
    pub language: LanguageDef,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct ConfJson {
    pub comments: Option<CommentDef>,
    pub brackets: Option<Vec<(String, String)>>,
}

pub fn load_language(name: &str) -> Option<LanguageDef> {
    // Handle running from both workspace root and crate root
    let p1 = Path::new("crates/code_editor/basic-languages").join(name).join(format!("{}.json", name));
    let p2 = Path::new("basic-languages").join(name).join(format!("{}.json", name));
    let path = if p1.exists() { p1 } else { p2 };
    if let Ok(content) = fs::read_to_string(path) {
        if let Ok(mut json) = serde_json::from_str::<LanguageJson>(&content) {
            if let Some(conf) = json.conf {
                json.language.comments = conf.comments;
                json.language.brackets = conf.brackets.unwrap_or_default();
            }
            return Some(json.language);
        }
    }
    None
}
