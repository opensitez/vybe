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
    Byte,
    SByte,
    Char,
    Short,
    UShort,
    Integer,
    UInteger,
    Long,
    ULong,
    Single,
    Double,
    Decimal,
    String,
    Boolean,
    Date,
    Variant,
    Object,
    Custom(String),
}

impl VBType {
    pub fn from_str(s: &str) -> Self {
        // Strip generic parameters (e.g. "List(Of String)" -> "List")
        let base_s = s.split('(').next().unwrap_or(s).trim();

        match base_s.to_lowercase().as_str() {
            "byte" => VBType::Byte,
            "sbyte" => VBType::SByte,
            "char" => VBType::Char,
            "short" | "int16" => VBType::Short,
            "ushort" | "uint16" => VBType::UShort,
            "integer" | "int32" => VBType::Integer,
            "uinteger" | "uint32" => VBType::UInteger,
            "long" | "int64" => VBType::Long,
            "ulong" | "uint64" => VBType::ULong,
            "single" => VBType::Single,
            "double" => VBType::Double,
            "decimal" => VBType::Decimal,
            "string" => VBType::String,
            "boolean" => VBType::Boolean,
            "date" | "datetime" => VBType::Date,
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
            VBType::Byte => write!(f, "Byte"),
            VBType::SByte => write!(f, "SByte"),
            VBType::Char => write!(f, "Char"),
            VBType::Short => write!(f, "Short"),
            VBType::UShort => write!(f, "UShort"),
            VBType::Integer => write!(f, "Integer"),
            VBType::UInteger => write!(f, "UInteger"),
            VBType::Long => write!(f, "Long"),
            VBType::ULong => write!(f, "ULong"),
            VBType::Single => write!(f, "Single"),
            VBType::Double => write!(f, "Double"),
            VBType::Decimal => write!(f, "Decimal"),
            VBType::String => write!(f, "String"),
            VBType::Boolean => write!(f, "Boolean"),
            VBType::Date => write!(f, "Date"),
            VBType::Variant => write!(f, "Variant"),
            VBType::Object => write!(f, "Object"),
            VBType::Custom(s) => write!(f, "{}", s),
        }
    }
}
