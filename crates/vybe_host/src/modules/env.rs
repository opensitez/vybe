use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // Get command line arguments as array of strings
    vm.register_host_fn("wasi:cli", "args", Box::new(|_args: &[Value]| {
        let args: Vec<Value> = std::env::args()
            .map(|a| Value::String(Rc::from(a.as_str())))
            .collect();
        Value::Object(Rc::new(RefCell::new(Object::new_array(args))))
    }));

    // Get environment variable (returns null if not set)
    vm.register_host_fn("wasi:cli", "getEnv", Box::new(|args: &[Value]| {
        let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        match std::env::var(&key) {
            Ok(val) => Value::String(Rc::from(val.as_str())),
            Err(_) => Value::Null,
        }
    }));

    // Get current working directory
    vm.register_host_fn("wasi:cli", "cwd", Box::new(|_args: &[Value]| {
        match std::env::current_dir() {
            Ok(p) => Value::String(Rc::from(p.to_string_lossy().as_ref())),
            Err(_) => Value::Null,
        }
    }));

    // Get platform name
    vm.register_host_fn("wasi:cli", "platform", Box::new(|_args: &[Value]| {
        Value::String(Rc::from(std::env::consts::OS))
    }));

    // Get architecture
    vm.register_host_fn("wasi:cli", "arch", Box::new(|_args: &[Value]| {
        Value::String(Rc::from(std::env::consts::ARCH))
    }));
}
