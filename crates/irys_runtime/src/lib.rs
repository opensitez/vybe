pub mod interpreter;
pub mod evaluator;
pub mod environment;
pub mod value;
pub mod event_system;
pub mod builtins;
pub mod file_io;
pub mod std_lib;
pub mod collections;

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
}

pub use interpreter::*;
pub use evaluator::*;
pub use environment::*;
pub use value::*;
pub use event_system::*;
