pub mod expr;
pub mod stmt;
pub mod decl;

pub use expr::*;
pub use stmt::*;
pub use decl::*;

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Program {
    pub declarations: Vec<Declaration>,
    pub statements: Vec<Statement>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Identifier(pub String);

impl Identifier {
    pub fn new(s: impl Into<String>) -> Self {
        Identifier(s.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum VBType {
    Integer,
    Long,
    Single,
    Double,
    String,
    Boolean,
    Variant,
    Object,
    Custom(String),
}

impl VBType {
    pub fn from_str(s: &str) -> Self {
        // Strip generic parameters (e.g. "List(Of String)" -> "List")
        let base_s = s.split('(').next().unwrap_or(s).trim();

        match base_s.to_lowercase().as_str() {
            "integer" => VBType::Integer,
            "long" => VBType::Long,
            "single" => VBType::Single,
            "double" => VBType::Double,
            "string" => VBType::String,
            "boolean" => VBType::Boolean,
            "variant" => VBType::Variant,
            "object" => VBType::Object,
            _ => VBType::Custom(base_s.to_string()),
        }
    }
}

use std::fmt;

impl fmt::Display for VBType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            VBType::Integer => write!(f, "Integer"),
            VBType::Long => write!(f, "Long"),
            VBType::Single => write!(f, "Single"),
            VBType::Double => write!(f, "Double"),
            VBType::String => write!(f, "String"),
            VBType::Boolean => write!(f, "Boolean"),
            VBType::Variant => write!(f, "Variant"),
            VBType::Object => write!(f, "Object"),
            VBType::Custom(s) => write!(f, "{}", s),
        }
    }
}
