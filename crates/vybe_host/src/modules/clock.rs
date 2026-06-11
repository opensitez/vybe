use std::sync::{Arc, Mutex, OnceLock};
use std::time::Instant;
use vybe_bytecode::value::Object;
use vybe_bytecode::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    // ── wasi:clocks/monotonic-clock — WASI 0.2 spec interface ───────────
    // Returns nanoseconds since an arbitrary reference point (process start).
    // Values are only meaningful relative to each other — use for scheduling.
    // Mirrors proposals/WASI/proposals/clocks/wit/monotonic-clock.wit.
    vm.register_host_fn(
        "wasi:clocks/monotonic-clock",
        "now",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            static START: std::sync::OnceLock<std::time::Instant> = std::sync::OnceLock::new();
            let start = START.get_or_init(std::time::Instant::now);
            Value::F64(start.elapsed().as_nanos() as f64)
        }),
    );

    vm.register_host_fn(
        "wasi:clocks/monotonic-clock",
        "resolution",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::F64(1.0) // 1 nanosecond resolution
        }),
    );

    // subscribe-instant(when: instant) → pollable
    // Returns a TimerPollable that becomes ready when monotonic clock reaches `when` ns.
    vm.register_host_fn(
        "wasi:clocks/monotonic-clock",
        "subscribe-instant",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ready_at_ns = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("TimerPollable")));
            obj.properties
                .insert("__ready_at_ns".into(), Value::F64(ready_at_ns));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // subscribe-duration(how-long: duration) → pollable
    // Returns a TimerPollable ready after `how-long` ns from now.
    vm.register_host_fn(
        "wasi:clocks/monotonic-clock",
        "subscribe-duration",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            static START: OnceLock<Instant> = OnceLock::new();
            let start = START.get_or_init(Instant::now);
            let how_long_ns = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let ready_at_ns = start.elapsed().as_nanos() as f64 + how_long_ns;
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("TimerPollable")));
            obj.properties
                .insert("__ready_at_ns".into(), Value::F64(ready_at_ns));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // ── WASI 0.3 additions ───────────────────────────────────────────
    // get-resolution — 0.3 rename of `resolution`; keep both for compat.
    vm.register_host_fn(
        "wasi:clocks/monotonic-clock",
        "get-resolution",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(1.0)),
    );

    // wait-until(when: mark) → future<result<_, error-code>>
    // Blocks until the monotonic clock reaches `when` ns, then returns a resolved future.
    vm.register_host_fn(
        "wasi:clocks/monotonic-clock",
        "wait-until",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            static START: OnceLock<Instant> = OnceLock::new();
            let start = START.get_or_init(Instant::now);
            let mark_ns = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let elapsed_ns = start.elapsed().as_nanos() as f64;
            if mark_ns > elapsed_ns {
                std::thread::sleep(std::time::Duration::from_nanos(
                    (mark_ns - elapsed_ns) as u64,
                ));
            }
            let (future_val, future_id) = ctx.create_future();
            ctx.resolve_future(future_id, Value::Null);
            future_val
        }),
    );

    // wait-for(how-long: duration) → future<result<_, error-code>>
    // Blocks for `how-long` ns, then returns a resolved future.
    vm.register_host_fn(
        "wasi:clocks/monotonic-clock",
        "wait-for",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let how_long_ns = args.first().map(|v| v.as_f64()).unwrap_or(0.0) as u64;
            if how_long_ns > 0 {
                std::thread::sleep(std::time::Duration::from_nanos(how_long_ns));
            }
            let (future_val, future_id) = ctx.create_future();
            ctx.resolve_future(future_id, Value::Null);
            future_val
        }),
    );

    // ── wasi:clocks/wall-clock — WASI 0.2.11 spec interface ─────────────
    // The canonical WASI wall-clock primitive. Returns a `datetime` record
    // `{ seconds: u64, nanoseconds: u32 }` per the .wit at
    // proposals/WASI/proposals/clocks/wit/wall-clock.wit. This is the
    // single source-of-truth timestamp; `ecma:date.now` reads through it.
    vm.register_host_fn(
        "wasi:clocks/wall-clock",
        "now",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let dur = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default();
            let mut rec = Object::new();
            rec.properties
                .insert("seconds".into(), Value::F64(dur.as_secs() as f64));
            rec.properties
                .insert("nanoseconds".into(), Value::F64(dur.subsec_nanos() as f64));
            Value::Object(Arc::new(Mutex::new(rec)))
        }),
    );

    // wasi:clocks/wall-clock.resolution — clock tick resolution per spec.
    // Most platforms report nanosecond resolution; we return 1ns.
    vm.register_host_fn(
        "wasi:clocks/wall-clock",
        "resolution",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut rec = Object::new();
            rec.properties.insert("seconds".into(), Value::F64(0.0));
            rec.properties.insert("nanoseconds".into(), Value::F64(1.0));
            Value::Object(Arc::new(Mutex::new(rec)))
        }),
    );

    // ── wasi:clocks/timezone — WASI 0.2 timezone interface ──────────────
    // display(when: datetime) → timezone-display { utc-offset: s32, name: string, in-daylight-saving-time: bool }
    vm.register_host_fn(
        "wasi:clocks/timezone",
        "display",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut rec = Object::new();
            rec.properties.insert("utc-offset".into(), Value::I32(0));
            rec.properties
                .insert("name".into(), Value::String(Arc::from("UTC")));
            rec.properties
                .insert("in-daylight-saving-time".into(), Value::Bool(false));
            Value::Object(Arc::new(Mutex::new(rec)))
        }),
    );

    // utc-offset(when: datetime) → s32  — seconds east of UTC
    vm.register_host_fn(
        "wasi:clocks/timezone",
        "utc-offset",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::I32(0)),
    );
}
