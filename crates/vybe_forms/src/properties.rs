use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum PropertyValue {
    String(String),
    Integer(i32),
    Boolean(bool),
    Double(f64),
    StringArray(Vec<String>),
    /// Raw code expression that should be written as-is (not quoted).
    Expression(String),
}

impl PropertyValue {
    pub fn as_string(&self) -> Option<&str> {
        match self {
            PropertyValue::String(s) => Some(s),
            _ => None,
        }
    }

    pub fn as_int(&self) -> Option<i32> {
        match self {
            PropertyValue::Integer(i) => Some(*i),
            _ => None,
        }
    }

    pub fn as_bool(&self) -> Option<bool> {
        match self {
            PropertyValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    pub fn as_double(&self) -> Option<f64> {
        match self {
            PropertyValue::Double(d) => Some(*d),
            _ => None,
        }
    }
    
    pub fn as_string_array(&self) -> Option<&Vec<String>> {
        match self {
            PropertyValue::StringArray(arr) => Some(arr),
            _ => None,
        }
    }
}

impl From<String> for PropertyValue {
    fn from(s: String) -> Self {
        PropertyValue::String(s)
    }
}

impl From<&str> for PropertyValue {
    fn from(s: &str) -> Self {
        PropertyValue::String(s.to_string())
    }
}

impl From<i32> for PropertyValue {
    fn from(i: i32) -> Self {
        PropertyValue::Integer(i)
    }
}

impl From<bool> for PropertyValue {
    fn from(b: bool) -> Self {
        PropertyValue::Boolean(b)
    }
}

impl From<f64> for PropertyValue {
    fn from(d: f64) -> Self {
        PropertyValue::Double(d)
    }
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct PropertyBag {
    properties: HashMap<String, PropertyValue>,
}

impl PropertyBag {
    pub fn new() -> Self {
        Self {
            properties: HashMap::new(),
        }
    }

    pub fn set(&mut self, key: impl Into<String>, value: impl Into<PropertyValue>) {
        self.properties.insert(key.into(), value.into());
    }
    
    pub fn set_raw(&mut self, key: impl Into<String>, value: PropertyValue) {
        self.properties.insert(key.into(), value);
    }

    pub fn get(&self, key: &str) -> Option<&PropertyValue> {
        self.properties.get(key)
    }

    pub fn get_string(&self, key: &str) -> Option<&str> {
        self.get(key).and_then(|v| v.as_string())
    }

    pub fn get_int(&self, key: &str) -> Option<i32> {
        self.get(key).and_then(|v| v.as_int())
    }

    pub fn get_bool(&self, key: &str) -> Option<bool> {
        self.get(key).and_then(|v| v.as_bool())
    }

    pub fn get_double(&self, key: &str) -> Option<f64> {
        self.get(key).and_then(|v| v.as_double())
    }
    
    pub fn get_string_array(&self, key: &str) -> Option<&Vec<String>> {
        self.get(key).and_then(|v| v.as_string_array())
    }

    pub fn iter(&self) -> impl Iterator<Item = (&String, &PropertyValue)> {
        self.properties.iter()
    }
}
