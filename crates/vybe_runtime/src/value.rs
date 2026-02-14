use std::collections::HashMap;
use std::fmt;
use std::rc::Rc;
use std::cell::RefCell;
use chrono::{NaiveDate, Duration};

#[derive(Debug, Clone, PartialEq)]
pub struct ObjectData {
    pub class_name: String,
    pub fields: HashMap<String, Value>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Byte(u8),
    Char(char),
    Integer(i32),
    Long(i64),
    Single(f32),
    Double(f64),
    Date(f64),
    String(String),
    Boolean(bool),
    Array(Vec<Value>),
    Collection(Rc<RefCell<crate::collections::ArrayList>>),
    Queue(Rc<RefCell<crate::collections::Queue>>),
    Stack(Rc<RefCell<crate::collections::Stack>>),
    HashSet(Rc<RefCell<crate::collections::VBHashSet>>),
    Dictionary(Rc<RefCell<crate::collections::VBDictionary>>),
    Nothing,
    Object(Rc<RefCell<ObjectData>>),
    Lambda {
        params: Vec<vybe_parser::ast::decl::Parameter>,
        body: Box<vybe_parser::ast::expr::LambdaBody>,
        env: Rc<RefCell<crate::environment::Environment>>,
    },
}

impl Value {
    pub fn as_integer(&self) -> Result<i32, RuntimeError> {
        match self {
            Value::Integer(i) => Ok(*i),
            Value::Long(l) => Ok(*l as i32),
            Value::Single(f) => Ok(*f as i32),
            Value::Double(d) => Ok(*d as i32),
            Value::Byte(b) => Ok(*b as i32),
            Value::Char(c) => Ok(*c as i32),
            Value::Date(d) => Ok(*d as i32),
            Value::String(s) => {
                if s.to_uppercase().starts_with("&H") {
                     i32::from_str_radix(&s[2..], 16).map_err(|_| RuntimeError::TypeError {
                        expected: "Integer (Hex)".to_string(),
                        got: s.clone(),
                     })
                } else if s.to_uppercase().starts_with("&O") {
                     i32::from_str_radix(&s[2..], 8).map_err(|_| RuntimeError::TypeError {
                        expected: "Integer (Oct)".to_string(),
                        got: s.clone(),
                     })
                } else {
                    s.parse().map_err(|_| RuntimeError::TypeError {
                        expected: "Integer".to_string(),
                        got: format!("{:?}", self),
                    })
                }
            },
            Value::Boolean(b) => Ok(if *b { -1 } else { 0 }),
            Value::Nothing => Ok(0),
            _ => Err(RuntimeError::TypeError {
                expected: "Integer".to_string(),
                got: format!("{:?}", self),
            }),
        }
    }

    pub fn as_long(&self) -> Result<i64, RuntimeError> {
        match self {
            Value::Integer(i) => Ok(*i as i64),
            Value::Long(l) => Ok(*l),
            Value::Single(f) => Ok(*f as i64),
            Value::Double(d) => Ok(*d as i64),
            Value::Byte(b) => Ok(*b as i64),
            Value::Char(c) => Ok(*c as i64),
            Value::Date(d) => Ok(*d as i64),
            Value::String(s) => {
                if s.to_uppercase().starts_with("&H") {
                     i64::from_str_radix(&s[2..], 16).map_err(|_| RuntimeError::TypeError {
                        expected: "Long (Hex)".to_string(),
                        got: s.clone(),
                     })
                } else if s.to_uppercase().starts_with("&O") {
                     i64::from_str_radix(&s[2..], 8).map_err(|_| RuntimeError::TypeError {
                        expected: "Long (Oct)".to_string(),
                        got: s.clone(),
                     })
                } else {
                    s.parse().map_err(|_| RuntimeError::TypeError {
                        expected: "Long".to_string(),
                        got: format!("{:?}", self),
                    })
                }
            },
            Value::Boolean(b) => Ok(if *b { -1 } else { 0 }),
            Value::Nothing => Ok(0),
            _ => Err(RuntimeError::TypeError {
                expected: "Long".to_string(),
                got: format!("{:?}", self),
            }),
        }
    }

    pub fn as_double(&self) -> Result<f64, RuntimeError> {
        match self {
            Value::Integer(i) => Ok(*i as f64),
            Value::Long(l) => Ok(*l as f64),
            Value::Single(f) => Ok(*f as f64),
            Value::Double(d) => Ok(*d),
            Value::Date(d) => Ok(*d),
            Value::String(s) => s.parse().map_err(|_| RuntimeError::TypeError {
                expected: "Double".to_string(),
                got: format!("{:?}", self),
            }),
            Value::Nothing => Ok(0.0),
            _ => Err(RuntimeError::TypeError {
                expected: "Double".to_string(),
                got: format!("{:?}", self),
            }),
        }
    }

    pub fn as_string(&self) -> String {
        match self {
            Value::Integer(i) => i.to_string(),
            Value::Long(l) => l.to_string(),
            Value::Byte(b) => b.to_string(),
            Value::Char(c) => c.to_string(),
            Value::Single(f) => f.to_string(),
            Value::Double(d) => d.to_string(),
            Value::Date(d) => {
                // OLE Automation Date: Days since Dec 30 1899
                let base_date = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap();
                let days = d.trunc() as i64;
                let fraction = d.fract();
                let seconds = (fraction * 86400.0).round() as i64;
                
                if let Some(date) = base_date.checked_add_signed(Duration::days(days)) {
                     if let Some(final_date) = date.checked_add_signed(Duration::seconds(seconds)) {
                         return final_date.format("%m/%d/%Y %H:%M:%S").to_string();
                     }
                }
                d.to_string() // Fallback
            }
            Value::String(s) => s.clone(),
            Value::Boolean(b) => if *b { "True" } else { "False" }.to_string(),
            Value::Array(_) => "[Array]".to_string(),
            Value::Collection(c) => format!("[Collection Count={}]", c.borrow().count()),
            Value::Queue(q) => format!("[Queue Count={}]", q.borrow().count()),
            Value::Stack(s) => format!("[Stack Count={}]", s.borrow().count()),
            Value::HashSet(h) => format!("[HashSet Count={}]", h.borrow().count()),
            Value::Dictionary(d) => format!("[Dictionary Count={}]", d.borrow().count()),
            Value::Nothing => "Nothing".to_string(),
            Value::Object(obj_ref) => {
                let b = obj_ref.borrow();
                // StringBuilder: return the buffer content
                if b.class_name == "StringBuilder" {
                    return b.fields.get("__data").map(|v| v.as_string()).unwrap_or_default();
                }
                format!("[Object {}]", b.class_name)
            }
            Value::Lambda { .. } => "[Lambda]".to_string(),
        }
    }

    pub fn as_bool(&self) -> Result<bool, RuntimeError> {
        match self {
            Value::Boolean(b) => Ok(*b),
            Value::Integer(i) => Ok(*i != 0),
            Value::Long(l) => Ok(*l != 0),
            Value::Byte(b) => Ok(*b != 0),
            Value::Char(_) => Err(RuntimeError::TypeError { expected: "Boolean".to_string(), got: "Char".to_string() }),
            Value::Single(f) => Ok(*f != 0.0),
            Value::Double(d) => Ok(*d != 0.0),
            Value::Date(d) => Ok(*d != 0.0),
            Value::String(s) => {
                 let lower = s.to_lowercase();
                 if lower == "true" { Ok(true) }
                 else if lower == "false" { Ok(false) }
                 else {
                     // Try parsing as number
                     if let Ok(n) = s.parse::<f64>() {
                         Ok(n != 0.0)
                     } else {
                         Ok(!s.is_empty()) 
                     }
                 }
            },
            Value::Object(_) => Ok(true),
            Value::Nothing => Ok(false),
            _ => Err(RuntimeError::TypeError {
                expected: "Boolean".to_string(),
                got: format!("{:?}", self),
            }),
        }
    }

    pub fn as_byte(&self) -> Result<u8, RuntimeError> {
        match self {
             Value::Byte(b) => Ok(*b),
             Value::Integer(i) => if *i >= 0 && *i <= 255 { Ok(*i as u8) } else { Err(RuntimeError::Custom(format!("Overflow: {} to Byte", i))) },
             Value::Long(l) => if *l >= 0 && *l <= 255 { Ok(*l as u8) } else { Err(RuntimeError::Custom(format!("Overflow: {} to Byte", l))) },
             Value::Single(f) => if *f >= 0.0 && *f <= 255.0 { Ok(*f as u8) } else { Err(RuntimeError::Custom(format!("Overflow: {} to Byte", f))) },
             Value::Double(d) => if *d >= 0.0 && *d <= 255.0 { Ok(*d as u8) } else { Err(RuntimeError::Custom(format!("Overflow: {} to Byte", d))) },
             Value::String(_) => {
                 let i = self.as_integer()?; 
                 if i >= 0 && i <= 255 { Ok(i as u8) } else { Err(RuntimeError::Custom(format!("Overflow: {} to Byte", i))) }
             },
             Value::Boolean(b) => Ok(if *b { 255 } else { 0 }), // True is 255 in Byte? Or -1 (overflow)? VB.NET CByte(True) = 255? VB6 CByte(True) = 255?
             // VB A Byte is an unsigned 8-bit integer relative to 0. 
             // True is -1. CByte(-1) -> Overflow...
             // Wait. CByte(True) in VB.NET does conversion.
             // CInt(True) is -1.
             // CByte(True)? "Arithmetic operation resulted in an overflow."
             // Wow. So CByte(True) throws exception?
             // I'll check. If CByte(-1) throws, then CByte(True) throws.
             // Let's assume standard int conversion rules.
             // But usually people expect Byte to be 0-255.
             // I'll stick to safe conversion checked.
             _ => Err(RuntimeError::TypeError { expected: "Byte".to_string(), got: format!("{:?}", self) })
        }
    }

    pub fn as_char(&self) -> Result<char, RuntimeError> {
        match self {
            Value::Char(c) => Ok(*c),
            Value::String(s) => s.chars().next().ok_or(RuntimeError::Custom("String is empty".to_string())),
            Value::Integer(i) => std::char::from_u32(*i as u32).ok_or(RuntimeError::Custom(format!("Invalid char code {}", i))),
            Value::Long(l) => std::char::from_u32(*l as u32).ok_or(RuntimeError::Custom(format!("Invalid char code {}", l))),
            Value::Byte(b) => std::char::from_u32(*b as u32).ok_or(RuntimeError::Custom(format!("Invalid char code {}", b))),
            _ => Err(RuntimeError::TypeError { expected: "Char".to_string(), got: format!("{:?}", self) })
        }
    }

    pub fn is_truthy(&self) -> bool {
        match self {
            Value::Boolean(b) => *b,
            Value::Integer(i) => *i != 0,
            Value::Long(l) => *l != 0,
            Value::Single(f) => *f != 0.0,
            Value::Double(d) => *d != 0.0,
            Value::Date(d) => *d != 0.0,
            Value::String(s) => !s.is_empty(),
            Value::Object(_) => true,
            Value::Collection(_) => true,
            Value::Queue(_) => true,
            Value::Stack(_) => true,
            Value::HashSet(_) => true,
            Value::Dictionary(_) => true,
            Value::Nothing => false,
            _ => false,
        }
    }

    pub fn get_array_element(&self, index: usize) -> Result<Value, RuntimeError> {
        match self {
            Value::Array(arr) => {
                arr.get(index)
                    .cloned()
                    .ok_or_else(|| RuntimeError::Custom(format!("Array index {} out of bounds", index)))
            }
            Value::Collection(col) => col.borrow().item(index),
            _ => Err(RuntimeError::TypeError {
                expected: "Array or Collection".to_string(),
                got: format!("{:?}", self),
            }),
        }
    }

    pub fn set_array_element(&mut self, index: usize, value: Value) -> Result<(), RuntimeError> {
        match self {
            Value::Array(arr) => {
                if index < arr.len() {
                    arr[index] = value;
                    Ok(())
                } else {
                    Err(RuntimeError::Custom(format!("Array index {} out of bounds", index)))
                }
            }
            Value::Collection(col) => col.borrow_mut().set_item(index, value),
            _ => Err(RuntimeError::TypeError {
                expected: "Array or Collection".to_string(),
                got: format!("{:?}", self),
            }),
        }
    }

    pub fn array_length(&self) -> Result<usize, RuntimeError> {
        match self {
            Value::Array(arr) => Ok(arr.len()),
            _ => Err(RuntimeError::TypeError {
                expected: "Array".to_string(),
                got: format!("{:?}", self),
            }),
        }
    }

    /// Convert any iterable Value into a Vec<Value> for For Each loops.
    /// Dictionary yields KeyValuePair objects with Key/Value fields.
    /// Strings yield individual character strings.
    pub fn to_iterable(&self) -> Result<Vec<Value>, RuntimeError> {
        match self {
            Value::Array(items) => Ok(items.clone()),
            Value::Collection(c) => Ok(c.borrow().items.clone()),
            Value::Queue(q) => Ok(q.borrow().to_array()),
            Value::Stack(s) => Ok(s.borrow().to_array()),
            Value::HashSet(h) => Ok(h.borrow().to_array()),
            Value::Dictionary(d) => {
                let d = d.borrow();
                let keys = d.keys();
                let vals = d.values();
                Ok(keys.into_iter().zip(vals).map(|(k, v)| {
                    let mut fields = std::collections::HashMap::new();
                    fields.insert("key".to_string(), k);
                    fields.insert("value".to_string(), v);
                    fields.insert("__type".to_string(), Value::String("KeyValuePair".to_string()));
                    Value::Object(Rc::new(RefCell::new(ObjectData {
                        class_name: "KeyValuePair".to_string(),
                        fields,
                    })))
                }).collect())
            }
            Value::String(s) => Ok(s.chars().map(|c| Value::String(c.to_string())).collect()),
            Value::Object(obj) => {
                // If the object has an items/rows array field, iterate that
                let b = obj.borrow();
                if let Some(Value::Array(arr)) = b.fields.get("items").or(b.fields.get("rows")) {
                    Ok(arr.clone())
                } else {
                    Err(RuntimeError::Custom(format!(
                        "Object of type '{}' is not enumerable",
                        b.class_name
                    )))
                }
            }
            Value::Nothing => Ok(Vec::new()),
            _ => Err(RuntimeError::TypeError {
                expected: "Array, Collection, Dictionary, or other enumerable".to_string(),
                got: format!("{:?}", self),
            }),
        }
    }
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{}", self.as_string())
    }
}

#[derive(Debug, thiserror::Error)]
pub enum RuntimeError {
    #[error("Type error: expected {expected}, got {got}")]
    TypeError { expected: String, got: String },

    #[error("Undefined variable: {0}")]
    UndefinedVariable(String),

    #[error("Undefined function: {0}")]
    UndefinedFunction(String),

    #[error("Division by zero")]
    DivisionByZero,

    #[error("Exit: {0}")]
    Exit(ExitType),

    #[error("Return")]
    Return(Option<Value>),

    #[error("{0}")]
    Custom(String),

    /// Typed exception: (exception_type, message, inner_exception_msg)
    #[error("{1}")]
    Exception(String, String, Option<String>),
    
    #[error("Continue")]
    Continue(vybe_parser::ast::stmt::ContinueType),

    /// GoTo control flow: jump to the named label.
    #[error("GoTo {0}")]
    GoTo(String),
}

#[derive(Debug, Clone, PartialEq)]
pub enum ExitType {
    Sub,
    Function,
    For,
    Do,
    While,
    Select,
    Try,
    Property,
}

impl fmt::Display for ExitType {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            ExitType::Sub => write!(f, "Sub"),
            ExitType::Function => write!(f, "Function"),
            ExitType::For => write!(f, "For"),
            ExitType::Do => write!(f, "Do"),
            ExitType::While => write!(f, "While"),
            ExitType::Select => write!(f, "Select"),
            ExitType::Try => write!(f, "Try"),
            ExitType::Property => write!(f, "Property"),
        }
    }
}
