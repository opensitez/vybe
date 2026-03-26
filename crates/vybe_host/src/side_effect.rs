use std::collections::{HashMap, VecDeque};
use vybe_bytecode::Value;

/// A cross-language property value for the UI layer.
/// Both VB's `Value` and the bytecode VM's `Value` convert into this.
/// The renderer only needs these primitive types.
#[derive(Debug, Clone, PartialEq)]
pub enum PropValue {
    Null,
    Bool(bool),
    Int(i64),
    Float(f64),
    String(String),
}

impl PropValue {
    pub fn as_string(&self) -> String {
        match self {
            PropValue::Null => String::new(),
            PropValue::Bool(b) => b.to_string(),
            PropValue::Int(n) => n.to_string(),
            PropValue::Float(n) => {
                if *n == (*n as i64) as f64 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            PropValue::String(s) => s.clone(),
        }
    }
}

impl std::fmt::Display for PropValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.as_string())
    }
}

/// Side effects that any language runtime can produce.
/// The UI renderer consumes these — it doesn't know which language produced them.
#[derive(Debug, Clone, PartialEq)]
pub enum SideEffect {
    ConsoleOutput(String),
    ConsoleClear,

    MsgBox {
        text: String,
        title: String,
    },

    PropertyChange {
        object: String,
        property: String,
        value: PropValue,
    },

    AddControl {
        form_name: String,
        control_name: String,
        control_type: String,
        left: i32,
        top: i32,
        width: i32,
        height: i32,
        parent_name: String,
    },

    Repaint {
        control_name: String,
    },

    FormClose {
        form_name: String,
    },

    FormShow {
        form_name: String,
    },

    FormShowDialog {
        form_name: String,
    },

    RunApplication {
        form_name: String,
    },

    InputBox {
        prompt: String,
        title: String,
        default_response: String,
    },

    DataSourceChanged {
        control_name: String,
        columns: Vec<String>,
        rows: Vec<Vec<String>>,
    },

    BindingPositionChanged {
        binding_source_name: String,
        position: i32,
        count: i32,
    },
}

/// Events from the UI back to the language runtime.
#[derive(Debug, Clone)]
pub enum UIEvent {
    Click { control: String },
    DblClick { control: String },
    TextChanged { control: String, text: String },
    KeyPress { control: String, key_char: char },
    KeyDown { control: String, key_code: i32, shift: bool, ctrl: bool, alt: bool },
    MouseDown { control: String, button: i32, x: i32, y: i32 },
    MouseUp { control: String, button: i32, x: i32, y: i32 },
    MouseMove { control: String, x: i32, y: i32 },
    SelectedIndexChanged { control: String, index: i32 },
    CheckedChanged { control: String, checked: bool },
    FormClosing { form: String },
    FormLoad { form: String },
    Timer { name: String },
    GotFocus { control: String },
    LostFocus { control: String },
}

/// Shared queue for side effects. Any runtime pushes, the UI drains.
/// Also holds event handler registrations for JS callbacks.
#[derive(Default)]
pub struct SideEffectQueue {
    effects: VecDeque<SideEffect>,
    /// Event handlers: key = "controlName.eventName" → VM callback Value
    pub event_handlers: HashMap<String, Value>,
}

impl std::fmt::Debug for SideEffectQueue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("SideEffectQueue")
            .field("effects_len", &self.effects.len())
            .field("event_handlers_len", &self.event_handlers.len())
            .finish()
    }
}

impl SideEffectQueue {
    pub fn new() -> Self {
        Self { effects: VecDeque::new(), event_handlers: HashMap::new() }
    }

    /// Register an event handler callback for a control+event pair.
    pub fn register_event(&mut self, control: &str, event: &str, callback: Value) {
        let key = format!("{}.{}", control, event);
        self.event_handlers.insert(key, callback);
    }

    /// Look up a registered event handler.
    pub fn get_event_handler(&self, control: &str, event: &str) -> Option<&Value> {
        let key = format!("{}.{}", control, event);
        self.event_handlers.get(&key)
    }

    pub fn push(&mut self, effect: SideEffect) {
        self.effects.push_back(effect);
    }

    pub fn drain(&mut self) -> Vec<SideEffect> {
        self.effects.drain(..).collect()
    }

    pub fn pop_front(&mut self) -> Option<SideEffect> {
        self.effects.pop_front()
    }

    pub fn is_empty(&self) -> bool {
        self.effects.is_empty()
    }
}
