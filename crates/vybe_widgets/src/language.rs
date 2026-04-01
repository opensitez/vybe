use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};
use std::env;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentDef {
    #[serde(rename = "lineComment")]
    pub line_comment: Option<String>,
    #[serde(rename = "blockComment")]
    pub block_comment: Option<(String, String)>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LanguageDef {
    pub keywords: HashSet<String>,
    pub type_keywords: HashSet<String>,
    pub constants: HashSet<String>,
    pub operators: HashSet<String>,
    pub ignore_case: bool,
    pub comments: Option<CommentDef>,
    pub brackets: Vec<(String, String)>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct LanguageJson {
    pub conf: Option<ConfJson>,
    pub language: serde_json::Map<String, serde_json::Value>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct ConfJson {
    pub comments: Option<CommentDef>,
    pub brackets: Option<Vec<(String, String)>>,
}

pub fn load_language(name: &str) -> Option<LanguageDef> {
    let name = name.to_lowercase();
    // Try multiple relative paths to find the basic-languages folder. Also
    // walk up from the executable location to locate the workspace root so
    // the binary can be run from `target/debug` or other CWDs.
    let mut candidates: Vec<PathBuf> = vec![
        PathBuf::from("crates/code_editor/basic-languages"),
        PathBuf::from("basic-languages"),
        PathBuf::from("../code_editor/basic-languages"),
    ];

    // Walk upward from current executable to find a workspace copy of the folder.
    if let Ok(exe) = env::current_exe() {
        let mut dir_opt = exe.parent();
        while let Some(dir) = dir_opt {
            let cand = dir.join("crates").join("code_editor").join("basic-languages");
            if cand.exists() {
                candidates.insert(0, cand);
                break;
            }
            dir_opt = dir.parent();
        }
    }

    for folder in &candidates {
        let path = folder.join(&name).join(format!("{}.json", &name));
        if let Ok(content) = fs::read_to_string(&path) {
            if let Ok(json) = serde_json::from_str::<LanguageJson>(&content) {
                let mut lang = LanguageDef { keywords: HashSet::new(), type_keywords: HashSet::new(), constants: HashSet::new(), operators: HashSet::new(), ignore_case: false, comments: None, brackets: Vec::new() };
                
                // Extract ignoreCase
                if let Some(ic) = json.language.get("ignoreCase").and_then(|v| v.as_bool()) { lang.ignore_case = ic; }
                
                // Harvest all keywords/functions/constants from various naming conventions
                for (k, v) in &json.language {
                    if let Some(arr) = v.as_array() {
                        let target_set = if k.to_lowercase().contains("keyword") { Some(&mut lang.keywords) }
                            else if k.to_lowercase().contains("function") { Some(&mut lang.keywords) }
                            else if k.to_lowercase().contains("constant") { Some(&mut lang.constants) }
                            else if k.to_lowercase().contains("operator") { Some(&mut lang.operators) }
                            else if k.to_lowercase().contains("type") { Some(&mut lang.type_keywords) }
                            else { None };
                        
                        if let Some(set) = target_set {
                            for val in arr {
                                if let Some(s) = val.as_str() { set.insert(s.to_string()); }
                            }
                        }
                    }
                }

                if let Some(conf) = json.conf {
                    lang.comments = conf.comments;
                    lang.brackets = conf.brackets.unwrap_or_default();
                }
                return Some(lang);
            }
        }
    }
    None
}
