use vybe_bytecode::{VM, Value, HostContext};
use std::collections::HashMap;

thread_local! {
    static TIMERS: std::cell::RefCell<HashMap<String, f64>> = std::cell::RefCell::new(HashMap::new());
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn("wasi:cli", "log", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        println!("{}", parts.join(" "));
        Value::Null
    }));

    vm.register_host_fn("wasi:cli", "error", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        eprintln!("{}", parts.join(" "));
        Value::Null
    }));

    vm.register_host_fn("wasi:cli", "warn", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        eprintln!("[warn] {}", parts.join(" "));
        Value::Null
    }));

    // stdin — read a line from standard input (blocking)
    vm.register_host_fn("wasi:cli", "readLine", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
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
            Ok(_) => Value::String(std::sync::Arc::from(line.trim_end_matches('\n').trim_end_matches('\r'))),
            Err(_) => Value::Null,
        }
    }));

    // console.time / console.timeEnd — simple profiling
    vm.register_host_fn("wasi:cli", "time", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let label = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "default".into());
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis() as f64;
        // Store in a thread-local map
        TIMERS.with(|t| t.borrow_mut().insert(label, now));
        Value::Null
    }));

    vm.register_host_fn("wasi:cli", "timeEnd", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let label = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "default".into());
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis() as f64;
        TIMERS.with(|t| {
            if let Some(start) = t.borrow_mut().remove(&label) {
                println!("{}: {}ms", label, now - start);
            }
        });
        Value::Null
    }));

    // exit — terminate with exit code
    vm.register_host_fn("wasi:cli", "exit", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let code = args.first().map(|v| v.as_f64() as i32).unwrap_or(0);
        std::process::exit(code);
    }));
}
