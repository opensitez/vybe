//! Vybe System Interface (VSI) modules.
//!
//! Each module registers host functions with (module, name) pairs on the VM,
//! following the WASI two-level namespace model.
//!
//! Modules:
//! - `vybe:console` — console I/O
//! - `vybe:math`    — math functions
//! - `vybe:string`  — string operations
//! - `vybe:array`   — array operations
//! - `vybe:convert` — type conversions
//! - `vybe:json`    — JSON serialization
//! - `vybe:fs`      — filesystem
//! - `vybe:clock`   — time and sleep
//! - `vybe:env`     — environment and args
//! - `vybe:random`  — random number generation
//! - `vybe:http`    — HTTP client
//! - `vybe:gui`     — GUI / form creation

pub mod console;
pub mod math;
pub mod string;
pub mod array;
pub mod convert;
pub mod json;
pub mod fs;
pub mod clock;
pub mod env;
pub mod random;
pub mod http;
pub mod object;
pub mod regex;
pub mod collections;
pub mod runtime;
pub mod database;
pub mod gui;
pub mod types;
pub mod sockets;
pub mod crypto;
pub mod xml;
pub mod threading;
pub mod data;
pub mod drawing;

use vybe_bytecode::VM;

/// Register all standard VSI modules on a VM (no GUI).
pub fn register_all(vm: &mut VM) {
    console::register(vm);
    math::register(vm);
    string::register(vm);
    array::register(vm);
    convert::register(vm);
    json::register(vm);
    fs::register(vm);
    clock::register(vm);
    env::register(vm);
    random::register(vm);
    http::register(vm);
    object::register(vm);
    regex::register(vm);
    collections::register(vm);
    runtime::register(vm);
    database::register(vm);
    types::register(vm);
    sockets::register(vm);
    crypto::register(vm);
    xml::register(vm);
    threading::register(vm);
    data::register(vm);
    drawing::register(vm);
    // Set up namespace objects and type method table AFTER all host functions are registered
    crate::namespaces::setup_namespaces(vm);
    crate::type_methods::register_all(vm);
}

/// Register all standard VSI modules + GUI module.
pub fn register_all_with_gui(
    vm: &mut VM,
    queue: std::rc::Rc<std::cell::RefCell<crate::SideEffectQueue>>,
) {
    register_all(vm);
    gui::register(vm, queue);
    // Re-setup namespaces to include GUI functions
    crate::namespaces::setup_namespaces(vm);
}
