use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub(crate) fn now_secs() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
}

pub fn register(vm: &mut VM) {
    // ── wasi:clocks/wall-clock — WASI 0.2.11 spec interface ─────────────
    // The canonical WASI wall-clock primitive. Returns a `datetime` record
    // `{ seconds: u64, nanoseconds: u32 }` per the .wit at
    // proposals/WASI/proposals/clocks/wit/wall-clock.wit. This is the
    // single source-of-truth timestamp; `ecma:date.now` reads through it.
    vm.register_host_fn("wasi:clocks/wall-clock", "now", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let dur = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default();
        let mut rec = Object::new();
        rec.properties.insert("seconds".into(), Value::F64(dur.as_secs() as f64));
        rec.properties.insert("nanoseconds".into(), Value::F64(dur.subsec_nanos() as f64));
        Value::Object(Arc::new(Mutex::new(rec)))
    }));

    // wasi:clocks/wall-clock.resolution — clock tick resolution per spec.
    // Most platforms report nanosecond resolution; we return 1ns.
    vm.register_host_fn("wasi:clocks/wall-clock", "resolution", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut rec = Object::new();
        rec.properties.insert("seconds".into(), Value::F64(0.0));
        rec.properties.insert("nanoseconds".into(), Value::F64(1.0));
        Value::Object(Arc::new(Mutex::new(rec)))
    }));

    // ── Legacy flat `wasi:clocks` namespace ─────────────────────────────
    // Pre-spec shape kept for backward compat with existing callers
    // (sleep, hrtime, stopwatch). New code should target
    // `wasi:clocks/wall-clock` / `wasi:clocks/monotonic-clock` per spec.

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

    // `taskDelay` retired — `Task.Delay(ms)` compiles to
    // `Op::THREAD_SPAWN` of a worker that calls `wasi:clocks/sleep(ms)`
    // (see emitter::threading::emit_task_delay). The Task object is
    // built by the VM's native THREAD_SPAWN handler. Pure WASM, zero
    // host fns.

    vm.register_host_fn("wasi:clocks", "stopwatchStart", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut inner = obj.lock().unwrap();
            let running = inner.properties.get("isrunning").map(|v| v.as_bool()).unwrap_or(false);
            if !running {
                inner.properties.insert("__start".into(), Value::F64(now_millis()));
                inner.properties.insert("isrunning".into(), Value::Bool(true));
            }
        }
        Value::Null
    }));

    vm.register_host_fn("wasi:clocks", "stopwatchStop", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut inner = obj.lock().unwrap();
            let running = inner.properties.get("isrunning").map(|v| v.as_bool()).unwrap_or(false);
            if running {
                let start = inner.properties.get("__start").map(|v| v.as_f64()).unwrap_or(0.0);
                let acc = inner.properties.get("__accumulated").map(|v| v.as_f64()).unwrap_or(0.0);
                inner.properties.insert("__accumulated".into(), Value::F64(acc + (now_millis() - start)));
                inner.properties.insert("isrunning".into(), Value::Bool(false));
            }
        }
        Value::Null
    }));

    vm.register_host_fn("wasi:clocks", "stopwatchReset", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut inner = obj.lock().unwrap();
            inner.properties.insert("__start".into(), Value::F64(0.0));
            inner.properties.insert("__accumulated".into(), Value::F64(0.0));
            inner.properties.insert("isrunning".into(), Value::Bool(false));
        }
        Value::Null
    }));

    vm.register_host_fn("wasi:clocks", "stopwatchElapsed", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let inner = obj.lock().unwrap();
            let acc = inner.properties.get("__accumulated").map(|v| v.as_f64()).unwrap_or(0.0);
            let running = inner.properties.get("isrunning").map(|v| v.as_bool()).unwrap_or(false);
            if running {
                let start = inner.properties.get("__start").map(|v| v.as_f64()).unwrap_or(0.0);
                return Value::F64(acc + (now_millis() - start));
            }
            return Value::F64(acc);
        }
        Value::F64(0.0)
    }));

    let elapsed_idx = *vm.host_registry.get(&("wasi:clocks".into(), "stopwatchElapsed".into())).unwrap();
    vm.register_host_fn("wasi:clocks", "stopwatchNew", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Stopwatch")));
        obj.properties.insert("__start".into(), Value::F64(0.0));
        obj.properties.insert("__accumulated".into(), Value::F64(0.0));
        obj.properties.insert("isrunning".into(), Value::Bool(false));

        let mut getter = Object::new();
        getter.kind = ObjectKind::HostFunction(elapsed_idx);
        let getter_val = Value::Object(Arc::new(Mutex::new(getter)));
        obj.properties.insert("__get_elapsedmilliseconds".into(), getter_val.clone());
        obj.properties.insert("__get_elapsed".into(), getter_val);
        Value::Object(Arc::new(Mutex::new(obj)))
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
        Value::String(Arc::from(format_timestamp(now_secs()).as_str()))
    }));

    // date/today → current date as string
    vm.register_host_fn("wasi:clocks", "vbDate", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let (y, m, d, _, _, _) = decompose_timestamp(now_secs());
        Value::String(Arc::from(format!("{:02}/{:02}/{:04}", m, d, y).as_str()))
    }));

    // time → current time as string
    vm.register_host_fn("wasi:clocks", "vbTime", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let (_, _, _, h, min, s) = decompose_timestamp(now_secs());
        Value::String(Arc::from(format!("{:02}:{:02}:{:02}", h, min, s).as_str()))
    }));

    // timer → seconds since midnight
    vm.register_host_fn("wasi:clocks", "vbTimer", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let time_of_day = now_secs() % 86400;
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
        _ => now_secs(),
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

fn now_millis() -> f64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis() as f64
}
