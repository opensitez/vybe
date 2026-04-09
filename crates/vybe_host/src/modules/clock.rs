use std::sync::Arc;
use vybe_bytecode::{VM, Value, HostContext};

pub fn register(vm: &mut VM) {
    // Returns milliseconds since Unix epoch (like JS Date.now())
    vm.register_host_fn("wasi:clocks", "now", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let ms = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis();
        Value::F64(ms as f64)
    }));

    // Returns nanoseconds (high-resolution timer, like performance.now())
    vm.register_host_fn("wasi:clocks", "hrtime", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let ns = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos();
        Value::F64(ns as f64)
    }));

    // Sleep for N milliseconds (blocking)
    vm.register_host_fn("wasi:clocks", "sleep", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64()).unwrap_or(0.0) as u64;
        std::thread::sleep(std::time::Duration::from_millis(ms));
        Value::Null
    }));

    // Format a timestamp as ISO 8601 string (simple implementation)
    vm.register_host_fn("wasi:clocks", "toISOString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
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

        Value::String(Arc::from(format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:03}Z",
            y, m + 1, remaining + 1, hours, minutes, seconds, millis
        ).as_str()))
    }));

    // setTimeout/setInterval are handled by the VM's set_timer opcode.
    // No host function needed — the compiler emits the opcode directly.

    // --- VB Date/Time functions ---

    // now() → current date/time as ISO string
    vm.register_host_fn("wasi:clocks", "vbNow", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        Value::String(Arc::from(format_timestamp(now).as_str()))
    }));

    // date/today → current date as string
    vm.register_host_fn("wasi:clocks", "vbDate", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        let (y, m, d, _, _, _) = decompose_timestamp(now);
        Value::String(Arc::from(format!("{:02}/{:02}/{:04}", m, d, y).as_str()))
    }));

    // time → current time as string
    vm.register_host_fn("wasi:clocks", "vbTime", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        let (_, _, _, h, min, s) = decompose_timestamp(now);
        Value::String(Arc::from(format!("{:02}:{:02}:{:02}", h, min, s).as_str()))
    }));

    // timer → seconds since midnight
    vm.register_host_fn("wasi:clocks", "vbTimer", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        let time_of_day = now % 86400;
        Value::F64(time_of_day as f64)
    }));

    // year(date) → year number
    vm.register_host_fn("wasi:clocks", "vbYear", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ts = parse_vb_date(args);
        let (y, _, _, _, _, _) = decompose_timestamp(ts);
        Value::F64(y as f64)
    }));

    // month(date) → month number
    vm.register_host_fn("wasi:clocks", "vbMonth", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ts = parse_vb_date(args);
        let (_, m, _, _, _, _) = decompose_timestamp(ts);
        Value::F64(m as f64)
    }));

    // day(date) → day number
    vm.register_host_fn("wasi:clocks", "vbDay", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ts = parse_vb_date(args);
        let (_, _, d, _, _, _) = decompose_timestamp(ts);
        Value::F64(d as f64)
    }));

    // hour(date) → hour
    vm.register_host_fn("wasi:clocks", "vbHour", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ts = parse_vb_date(args);
        let (_, _, _, h, _, _) = decompose_timestamp(ts);
        Value::F64(h as f64)
    }));

    // minute(date) → minute
    vm.register_host_fn("wasi:clocks", "vbMinute", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ts = parse_vb_date(args);
        let (_, _, _, _, min, _) = decompose_timestamp(ts);
        Value::F64(min as f64)
    }));

    // second(date) → second
    vm.register_host_fn("wasi:clocks", "vbSecond", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ts = parse_vb_date(args);
        let (_, _, _, _, _, s) = decompose_timestamp(ts);
        Value::F64(s as f64)
    }));
}

fn parse_vb_date(args: &[Value]) -> u64 {
    // For now, treat numeric as epoch seconds, string as current time
    match args.first() {
        Some(Value::F64(n)) => *n as u64,
        _ => std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs(),
    }
}

fn decompose_timestamp(total_secs: u64) -> (i64, u64, u64, u64, u64, u64) {
    let days = total_secs / 86400;
    let time_of_day = total_secs % 86400;
    let hours = time_of_day / 3600;
    let minutes = (time_of_day % 3600) / 60;
    let seconds = time_of_day % 60;

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
    let mut m: u64 = 1;
    for md in &month_days {
        if remaining < *md { break; }
        remaining -= md;
        m += 1;
    }
    (y, m, (remaining + 1) as u64, hours, minutes, seconds)
}

fn format_timestamp(total_secs: u64) -> String {
    let (y, m, d, h, min, s) = decompose_timestamp(total_secs);
    format!("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", y, m, d, h, min, s)
}
