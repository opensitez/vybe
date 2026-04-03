use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // Get command line arguments as array of strings
    vm.register_host_fn("wasi:cli", "args", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let args: Vec<Value> = std::env::args()
            .map(|a| Value::String(Rc::from(a.as_str())))
            .collect();
        Value::Object(Rc::new(RefCell::new(Object::new_array(args))))
    }));

    // Get environment variable (returns null if not set)
    vm.register_host_fn("wasi:cli", "getEnv", Box::new(|_vm: &mut VM, args: &[Value]| {
        let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        match std::env::var(&key) {
            Ok(val) => Value::String(Rc::from(val.as_str())),
            Err(_) => Value::Null,
        }
    }));

    // Get current working directory
    vm.register_host_fn("wasi:cli", "cwd", Box::new(|_vm: &mut VM, _args: &[Value]| {
        match std::env::current_dir() {
            Ok(p) => Value::String(Rc::from(p.to_string_lossy().as_ref())),
            Err(_) => Value::Null,
        }
    }));

    // Get platform name
    vm.register_host_fn("wasi:cli", "platform", Box::new(|_vm: &mut VM, _args: &[Value]| {
        Value::String(Rc::from(std::env::consts::OS))
    }));

    // Get architecture
    vm.register_host_fn("wasi:cli", "arch", Box::new(|_vm: &mut VM, _args: &[Value]| {
        Value::String(Rc::from(std::env::consts::ARCH))
    }));

    // Machine name
    vm.register_host_fn("wasi:cli", "machineName", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let name = std::env::var("HOSTNAME")
            .or_else(|_| std::env::var("COMPUTERNAME"))
            .unwrap_or_else(|_| "unknown".into());
        Value::String(Rc::from(name.as_str()))
    }));

    // User name
    vm.register_host_fn("wasi:cli", "userName", Box::new(|_vm: &mut VM, _args: &[Value]| {
        Value::String(Rc::from(std::env::var("USER").or_else(|_| std::env::var("USERNAME")).unwrap_or_else(|_| "unknown".into()).as_str()))
    }));

    // Newline
    vm.register_host_fn("wasi:cli", "newLine", Box::new(|_vm: &mut VM, _args: &[Value]| {
        Value::String(Rc::from("\n"))
    }));

    // Tick count (ms since boot, approximated as ms since epoch)
    vm.register_host_fn("wasi:cli", "tickCount", Box::new(|_vm: &mut VM, _args: &[Value]| {
        let ms = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis();
        Value::F64(ms as f64)
    }));
}
