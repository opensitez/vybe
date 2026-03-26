use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct Identifier(pub String);

impl Identifier {
    pub fn new<S: Into<String>>(s: S) -> self::Identifier {
        let val: String = s.into();
        // Strip VB.NET escaped identifier brackets: [Stop] → Stop
        if val.starts_with('[') && val.ends_with(']') && val.len() > 2 {
            Identifier(val[1..val.len()-1].to_string())
        } else {
            Identifier(val)
        }
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl std::fmt::Display for Identifier {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.0)
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
    Date,
    Object,
    Variant,
    Custom(String),
}

impl VBType {
    pub fn from_str(s: &str) -> Self {
        match s.to_lowercase().as_str() {
            "integer" | "int32" => VBType::Integer,
            "long" | "int64" => VBType::Long,
            "single" | "float" => VBType::Single,
            "double" => VBType::Double,
            "string" => VBType::String,
            "boolean" | "bool" => VBType::Boolean,
            "date" | "datetime" => VBType::Date,
            "object" => VBType::Object,
            "variant" => VBType::Variant,
            _ => VBType::Custom(s.to_string()),
        }
    }
}

impl std::fmt::Display for VBType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            VBType::Integer => write!(f, "Integer"),
            VBType::Long => write!(f, "Long"),
            VBType::Single => write!(f, "Single"),
            VBType::Double => write!(f, "Double"),
            VBType::String => write!(f, "String"),
            VBType::Boolean => write!(f, "Boolean"),
            VBType::Date => write!(f, "Date"),
            VBType::Object => write!(f, "Object"),
            VBType::Variant => write!(f, "Variant"),
            VBType::Custom(s) => write!(f, "{}", s),
        }
    }
}
