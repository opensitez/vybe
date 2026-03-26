use std::rc::Rc;
use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    // Returns milliseconds since Unix epoch (like JS Date.now())
    vm.register_host_fn("wasi:clocks", "now", Box::new(|_args: &[Value]| {
        let ms = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis();
        Value::F64(ms as f64)
    }));

    // Returns nanoseconds (high-resolution timer, like performance.now())
    vm.register_host_fn("wasi:clocks", "hrtime", Box::new(|_args: &[Value]| {
        let ns = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos();
        Value::F64(ns as f64)
    }));

    // Sleep for N milliseconds (blocking)
    vm.register_host_fn("wasi:clocks", "sleep", Box::new(|args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64()).unwrap_or(0.0) as u64;
        std::thread::sleep(std::time::Duration::from_millis(ms));
        Value::Null
    }));

    // Format a timestamp as ISO 8601 string (simple implementation)
    vm.register_host_fn("wasi:clocks", "toISOString", Box::new(|args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64()).unwrap_or(0.0) as u64;
        let total_secs = ms / 1000;
        let millis = ms % 1000;

        // Simple date calculation from epoch seconds
        let days = total_secs / 86400;
        let time_of_day = total_secs % 86400;
        let hours = time_of_day / 3600;
        let minutes = (time_of_day % 3600) / 60;
        let seconds = time_of_day % 60;

        // Days since 1970-01-01 to Y-M-D (simplified, no leap second)
        let mut y = 1970i64;
        let mut remaining = days as i64;
        loop {
            let days_in_year = if y % 4 == 0 && (y % 100 != 0 || y % 400 == 0) { 366 } else { 365 };
            if remaining < days_in_year { break; }
            remaining -= days_in_year;
            y += 1;
        }
        let leap = y % 4 == 0 && (y % 100 != 0 || y % 400 == 0);
        let month_days = [31, if leap { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        let mut m = 0;
        for md in &month_days {
            if remaining < *md { break; }
            remaining -= md;
            m += 1;
        }

        Value::String(Rc::from(format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:03}Z",
            y, m + 1, remaining + 1, hours, minutes, seconds, millis
        ).as_str()))
    }));

    // setTimeout/setInterval are handled by the VM's set_timer opcode.
    // No host function needed — the compiler emits the opcode directly.
}
