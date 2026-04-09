//! Namespace objects — sets up VM globals as nested objects
//! that any language can traverse via struct_get chains.
//!
//! Each sub-module registers one domain (math, io, forms, etc.).

mod math;
mod console;
mod string;
mod convert;
mod array;
mod io;
mod environment;
mod forms;
mod gui;
mod net;
mod json;
mod types;
mod threading;
mod data;

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{VM, Value, HostContext};

/// Register all namespace objects on the VM.
/// Call this AFTER registering all host functions.
pub fn setup_namespaces(vm: &mut VM) {
    math::register(vm);
    console::register(vm);
    string::register(vm);
    convert::register(vm);
    array::register(vm);
    io::register(vm);
    environment::register(vm);
    forms::register(vm);
    gui::register(vm);
    net::register(vm);
    json::register(vm);
    types::register(vm);
    threading::register(vm);
    data::register(vm);
}

// ---- Shared helpers used by all sub-modules ----

/// Ensure a namespace object exists at the given dotted path.
/// Creates intermediate objects as needed. Returns the leaf object.
pub(crate) fn ensure_namespace(vm: &mut VM, path: &[&str]) -> Value {
    if path.is_empty() { return Value::Null; }
    let root_key = path[0].to_lowercase();
    let root = if let Some(existing) = vm.globals.get(&root_key) {
        existing.clone()
    } else {
        let obj = Value::Object(Arc::new(Mutex::new(Object::new())));
        vm.globals.insert(root_key.clone(), obj.clone());
        obj
    };
    let mut current = root;
    for &segment in &path[1..] {
        let key = segment.to_lowercase();
        let next = if let Value::Object(ref obj) = current {
            obj.lock().unwrap().properties.get(&key).cloned()
        } else {
            None
        };
        if let Some(existing) = next {
            current = existing;
        } else {
            let new_obj = Value::Object(Arc::new(Mutex::new(Object::new())));
            if let Value::Object(ref obj) = current {
                obj.lock().unwrap().properties.insert(key, new_obj.clone());
            }
            current = new_obj;
        }
    }
    current
}

/// Set a property on a namespace object (lowercased key).
pub(crate) fn set_prop(ns: &Value, name: &str, value: Value) {
    if let Value::Object(obj) = ns {
        obj.lock().unwrap().properties.insert(name.to_lowercase(), value);
    }
}

/// Create a HostFunction Value referencing a registered host function.
pub(crate) fn host_fn_ref(vm: &VM, module: &str, name: &str) -> Value {
    if let Some(&idx) = vm.host_registry.get(&(module.to_string(), name.to_string())) {
        let mut obj = Object::new();
        obj.properties.insert("__host_module".into(), Value::String(Arc::from(module)));
        obj.properties.insert("__host_name".into(), Value::String(Arc::from(name)));
        obj.properties.insert("__host_idx".into(), Value::F64(idx as f64));
        obj.kind = ObjectKind::HostFunction(idx);
        Value::Object(Arc::new(Mutex::new(obj)))
    } else {
        Value::Null
    }
}
