use crate::ast::expr::Expression;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum XmlNode {
    Element(XmlElement),
    Text(String),
    EmbeddedExpression(Expression),
    Comment(String),
    CData(String),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct XmlElement {
    pub name: XmlName,
    pub attributes: Vec<XmlAttribute>,
    pub children: Vec<XmlNode>,
    pub is_empty: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct XmlName {
    pub local: String,
    pub prefix: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct XmlAttribute {
    pub name: XmlName,
    pub value: Vec<XmlNode>, // Attributes can contain embedded expressions too: name="<%= expr %>"
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct XmlDocument {
    pub root: XmlElement,
    pub declaration: Option<XmlDeclaration>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct XmlDeclaration {
    pub version: String,
    pub encoding: Option<String>,
    pub standalone: Option<String>,
}
