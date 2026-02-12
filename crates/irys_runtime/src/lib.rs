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

pub use interpreter::*;
pub use evaluator::*;
pub use environment::*;
pub use value::*;
pub use event_system::*;
