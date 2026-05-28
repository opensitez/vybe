//! `node:readline` — Node.js readline module.
//!
//! Reference: <https://nodejs.org/api/readline.html>.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::VM;

fn make_interface() -> Value {
    let mut o = Object::new();
    // Methods — stub Undefined values so property-existence checks pass
    for m in ["close","pause","resume","setPrompt","getPrompt","prompt",
              "question","write","getCursorPos",
              "on","once","off","emit","removeListener","addListener","removeAllListeners"] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    // Properties
    o.properties.insert("terminal".into(), Value::Bool(false));
    o.properties.insert("line".into(), Value::String(Arc::from("")));
    o.properties.insert("cursor".into(), Value::I32(0));
    Value::Object(Arc::new(Mutex::new(o)))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn("node:readline", "createInterface", Box::new(|_ctx, _args| {
        make_interface()
    }));

    vm.register_host_fn("node:readline", "Interface", Box::new(|_ctx, _args| {
        make_interface()
    }));

    vm.register_host_fn("node:readline", "cursorTo", Box::new(|_ctx, _args| Value::Undefined));
    vm.register_host_fn("node:readline", "moveCursor", Box::new(|_ctx, _args| Value::Undefined));
    vm.register_host_fn("node:readline", "clearLine", Box::new(|_ctx, _args| Value::Undefined));
    vm.register_host_fn("node:readline", "clearScreenDown", Box::new(|_ctx, _args| Value::Undefined));
    vm.register_host_fn("node:readline", "emitKeypressEvents", Box::new(|_ctx, _args| Value::Undefined));
}
