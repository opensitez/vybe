use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("wasi:cli", "log", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        println!("{}", parts.join(" "));
        Value::Null
    }));

    vm.register_host_fn("wasi:cli", "error", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        eprintln!("{}", parts.join(" "));
        Value::Null
    }));

    vm.register_host_fn("wasi:cli", "warn", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        eprintln!("[warn] {}", parts.join(" "));
        Value::Null
    }));

    // stdin — read a line from standard input (blocking)
    vm.register_host_fn("wasi:cli", "readLine", Box::new(|args: &[Value]| {
        // Optional prompt
        if let Some(prompt) = args.first() {
            let p = format!("{}", prompt);
            if !p.is_empty() {
                use std::io::Write;
                print!("{}", p);
                let _ = std::io::stdout().flush();
            }
        }
        let mut line = String::new();
        match std::io::stdin().read_line(&mut line) {
            Ok(_) => Value::String(std::rc::Rc::from(line.trim_end_matches('\n').trim_end_matches('\r'))),
            Err(_) => Value::Null,
        }
    }));

    // exit — terminate with exit code
    vm.register_host_fn("wasi:cli", "exit", Box::new(|args: &[Value]| {
        let code = args.first().map(|v| v.as_f64() as i32).unwrap_or(0);
        std::process::exit(code);
    }));
}
