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
    },
}

pub use interpreter::*;
pub use evaluator::*;
pub use environment::*;
pub use value::*;
pub use event_system::*;
