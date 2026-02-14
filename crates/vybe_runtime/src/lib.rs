pub mod interpreter;
pub mod evaluator;
pub mod environment;
pub mod value;
pub mod event_system;
pub mod builtins;
pub mod file_io;
pub mod std_lib;
pub mod collections;
pub mod data_access;

/// A resource entry passed from the project layer into the runtime.
/// Carries type info so the runtime can distinguish strings from file resources.
#[derive(Debug, Clone, PartialEq)]
pub struct ResourceEntry {
    pub name: String,
    pub value: String,
    /// "string", "image", "icon", "audio", "file", "other"
    pub resource_type: String,
    /// For file-based resources: the resolved file path on disk
    pub file_path: Option<String>,
}

impl ResourceEntry {
    pub fn string(name: impl Into<String>, value: impl Into<String>) -> Self {
        Self { name: name.into(), value: value.into(), resource_type: "string".into(), file_path: None }
    }
    pub fn file(name: impl Into<String>, path: impl Into<String>, resource_type: impl Into<String>) -> Self {
        let p: String = path.into();
        Self { name: name.into(), value: p.clone(), resource_type: resource_type.into(), file_path: Some(p) }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RuntimeSideEffect {
    MsgBox(String),
    PropertyChange {
        object: String,
        property: String,
        value: Value,
    },
    ConsoleOutput(String),
    ConsoleClear,
    /// Signals that a data-bound control's data source has changed and needs re-rendering.
    DataSourceChanged {
        control_name: String,
        columns: Vec<String>,
        rows: Vec<Vec<String>>,
    },
    /// Signals BindingSource position change — bound controls should refresh.
    BindingPositionChanged {
        binding_source_name: String,
        position: i32,
        count: i32,
    },
    /// Close a form (fires FormClosing then FormClosed).
    FormClose {
        form_name: String,
    },
    /// Show a user form as a modal dialog.
    FormShowDialog {
        form_name: String,
    },
    /// Dynamically add a control to a form at runtime.
    AddControl {
        form_name: String,
        control_name: String,
        control_type: String,
        left: i32,
        top: i32,
        width: i32,
        height: i32,
    },
    /// Show a native InputBox dialog.
    InputBox {
        prompt: String,
        title: String,
        default_response: String,
        x_pos: i32,
        y_pos: i32,
    },
    /// Start the application message loop with the given form.
    RunApplication {
        form_name: String,
    },
}

// ---------------------------------------------------------------------------
// Thread-safe console I/O channel types (used for interactive console apps)
// ---------------------------------------------------------------------------

/// Messages sent from the interpreter thread to the UI.
#[derive(Debug, Clone)]
pub enum ConsoleMessage {
    /// Console output with color information (fg/bg use .NET ConsoleColor values 0-15).
    Output { text: String, fg: i32, bg: i32 },
    Clear,
    /// The interpreter is waiting for user input (Console.ReadLine).
    InputRequest,
    /// Sub Main finished successfully.
    Finished,
    /// Sub Main ended with an error.
    Error(String),
}

/// Event data passed from the UI layer to the interpreter for populating EventArgs fields.
#[derive(Debug, Clone)]
pub enum EventData {
    /// Mouse event data: button (0x100000=Left, 0x200000=Right, 0x400000=Middle), clicks, x, y, scroll delta
    Mouse { button: i32, clicks: i32, x: i32, y: i32, delta: i32 },
    /// Key down/up event data
    Key { key_code: i32, shift: bool, ctrl: bool, alt: bool },
    /// KeyPress event data (character input)
    KeyPress { key_char: char },
}

pub use interpreter::*;
pub use evaluator::*;
pub use environment::*;
pub use value::*;
pub use event_system::*;
