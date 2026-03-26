use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:console", "log", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        println!("{}", parts.join(" "));
        Value::Null
    }));

    vm.register_host_fn("vybe:console", "error", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        eprintln!("{}", parts.join(" "));
        Value::Null
    }));

    vm.register_host_fn("vybe:console", "warn", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        eprintln!("[warn] {}", parts.join(" "));
        Value::Null
    }));
}
