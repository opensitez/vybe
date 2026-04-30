//! `ecma:date` — ECMA-262 §21.4 Date mirror + cross-language adapters.
//!
//! **Not in any WebAssembly CG proposal.** Named `ecma:*` per the
//! project convention: ECMA-262 runtime shapes that have no merged WASM
//! proposal live under `ecma:*`. Only `wasm:js-string` and the
//! stage-1 `wasm:js-{number,boolean,undefined,symbol,bigint}` are real
//! WebAssembly names. See `JS_BUILTIN_CONVENTIONS.md`.
//!
//! Every function below operates on **milliseconds since Unix epoch**
//! (the ECMA-262 Date internal `[[DateValue]]` representation). The
//! `fromUnixSeconds` / `toUnixSeconds` helpers bridge to POSIX-style
//! seconds for PHP / C / Unix-shell workflows.
//!
//! All date math is delegated to `chrono::DateTime<Utc>` so we don't
//! reinvent leap-year / timezone arithmetic.

use std::sync::Arc;
use chrono::{DateTime, Datelike, NaiveDate, NaiveDateTime, NaiveTime, TimeZone, Timelike, Utc};
use vybe_bytecode::{HostContext, Value, VM};

const MODULE: &str = "ecma:date";

fn dt_from_ms(ms: f64) -> Option<DateTime<Utc>> {
    let secs = (ms / 1000.0).floor() as i64;
    let nsecs = (((ms.rem_euclid(1000.0)) * 1_000_000.0) as u32).min(999_999_999);
    Utc.timestamp_opt(secs, nsecs).single()
}

fn ms_of(dt: DateTime<Utc>) -> f64 {
    (dt.timestamp() as f64) * 1000.0 + (dt.timestamp_subsec_millis() as f64)
}

/// Extract the millisecond timestamp from a Date arg. Accepts either a
/// raw number (legacy phpDate / direct host call form) or a Date Object
/// with `__time` property (the canonical instance shape produced by
/// `ecma:date.new`). Defaults to NaN when neither applies so downstream
/// `dt_from_ms` returns `Invalid Date`.
fn ms_arg(args: &[Value], idx: usize) -> f64 {
    match args.get(idx) {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            match o.properties.get("__time") {
                Some(v) => v.as_f64(),
                None => f64::NAN,
            }
        }
        Some(v) => v.as_f64(),
        None => f64::NAN,
    }
}

/// Apply a single setter (`setFullYear`, `setHours`, ...) to the Date in
/// `args[0]` using the new component value at `args[1]`. Returns the
/// resulting ms timestamp. `component` is the chrono field name.
///
/// Overflow handling matches ECMA-262 §21.4.1.13 (MakeDay/MakeTime):
/// out-of-range component values roll into adjacent components rather
/// than producing NaN. e.g. `setDate(35)` on Jan rolls into Feb.
fn setter_helper(args: &[Value], component: &str) -> f64 {
    let ms = ms_arg(args, 0);
    let val = args.get(1).map(|v| v.as_f64()).unwrap_or(f64::NAN);
    if val.is_nan() { return f64::NAN; }
    let dt = match dt_from_ms(ms) { Some(d) => d, None => return f64::NAN };

    // For day/month/year we rebuild via days-since-epoch arithmetic so
    // out-of-range values overflow per spec. For h/m/s/ms we offset the
    // ms timestamp directly — same overflow semantics for free.
    match component {
        "year" => {
            // setFullYear keeps month/day from current dt but allows year.
            // Use overflow-friendly reconstruction.
            let y = val as i32;
            let m = dt.month();
            let d = dt.day();
            match NaiveDate::from_ymd_opt(y, m, 1) {
                Some(start) => {
                    let new_date = start + chrono::Duration::days((d - 1) as i64);
                    let new_time = dt.time();
                    ms_of(Utc.from_utc_datetime(&new_date.and_time(new_time)))
                }
                None => f64::NAN,
            }
        }
        "month" => {
            // setMonth(11) on year=2024 day=31 → 2024-12-31, no overflow
            // setMonth(12) on year=2024 day=31 → roll to 2025-01-31
            let y = dt.year();
            let m_zero = val as i64; // 0-indexed month requested
            let total_months = (y as i64) * 12 + m_zero;
            let new_y = total_months.div_euclid(12) as i32;
            let new_m = (total_months.rem_euclid(12) + 1) as u32;
            let d = dt.day();
            // First of new month, then add (d-1) days for overflow rollover
            match NaiveDate::from_ymd_opt(new_y, new_m, 1) {
                Some(start) => {
                    let new_date = start + chrono::Duration::days((d - 1) as i64);
                    let new_time = dt.time();
                    ms_of(Utc.from_utc_datetime(&new_date.and_time(new_time)))
                }
                None => f64::NAN,
            }
        }
        "day" => {
            // setDate(n): replace day-of-month, allowing overflow into
            // later months. Compute as: first-of-month + (n-1) days.
            let y = dt.year();
            let m = dt.month();
            let d = val as i64;
            match NaiveDate::from_ymd_opt(y, m, 1) {
                Some(start) => {
                    let new_date = start + chrono::Duration::days(d - 1);
                    let new_time = dt.time();
                    ms_of(Utc.from_utc_datetime(&new_date.and_time(new_time)))
                }
                None => f64::NAN,
            }
        }
        "hour" => {
            // ms_offset = (new_h - cur_h) * 3600_000
            let cur_h = dt.hour() as f64;
            ms + (val - cur_h) * 3_600_000.0
        }
        "minute" => {
            let cur_min = dt.minute() as f64;
            ms + (val - cur_min) * 60_000.0
        }
        "second" => {
            let cur_s = dt.second() as f64;
            ms + (val - cur_s) * 1000.0
        }
        "millisecond" => {
            let cur_ms = dt.timestamp_subsec_millis() as f64;
            ms + (val - cur_ms)
        }
        _ => f64::NAN,
    }
}

/// Format ms-since-epoch as an ISO 8601 string (no surrounding quotes).
/// Returns None for NaN / out-of-range values. Used by JSON.stringify to
/// serialize Date instances without re-locking the object.
pub fn format_iso_from_ms(ms: f64) -> Option<String> {
    dt_from_ms(ms).map(|dt| dt.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string())
}

/// Dispatch a Date method by name. Used by `ecma:value.invokeMethod` when
/// the polymorphic shim sees a `__type=Date` Object — the registered host
/// fns are reached normally for `d.method()` direct calls (typeregistry
/// vtable), but `invokeMethod` doesn't carry receiver type info, so the
/// shim re-resolves here.
///
/// Returns `Some(result)` for known method names; `None` so the caller
/// can fall through to "method not found" (Undefined).
pub fn dispatch_date_method(method: &str, args: &[Value]) -> Option<Value> {
    // Getters: read `__time` from args[0], compute, return.
    macro_rules! getter {
        ($body:expr) => {{
            let ms = ms_arg(args, 0);
            if let Some(dt) = dt_from_ms(ms) {
                Value::F64($body(&dt) as f64)
            } else {
                Value::F64(f64::NAN)
            }
        }};
    }
    let result = match method {
        "getFullYear" | "getUTCFullYear" => getter!(|dt: &DateTime<Utc>| dt.year()),
        "getMonth" | "getUTCMonth" => getter!(|dt: &DateTime<Utc>| dt.month() as i32 - 1),
        "getDate" | "getUTCDate" => getter!(|dt: &DateTime<Utc>| dt.day() as i32),
        "getDay" | "getUTCDay" => getter!(|dt: &DateTime<Utc>| dt.weekday().num_days_from_sunday() as i32),
        "getHours" | "getUTCHours" => getter!(|dt: &DateTime<Utc>| dt.hour() as i32),
        "getMinutes" | "getUTCMinutes" => getter!(|dt: &DateTime<Utc>| dt.minute() as i32),
        "getSeconds" | "getUTCSeconds" => getter!(|dt: &DateTime<Utc>| dt.second() as i32),
        "getMilliseconds" | "getUTCMilliseconds" => getter!(|dt: &DateTime<Utc>| dt.timestamp_subsec_millis() as i32),
        "getTime" | "valueOf" => Value::F64(ms_arg(args, 0)),
        "getTimezoneOffset" => Value::F64(0.0),
        "toISOString" | "toJSON" => {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string().as_str()
                )),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        "toString" => {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%a %b %d %Y %H:%M:%S GMT+0000 (UTC)").to_string().as_str()
                )),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        "toDateString" => {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%a %b %d %Y").to_string().as_str()
                )),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        "toTimeString" => {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%H:%M:%S GMT+0000 (UTC)").to_string().as_str()
                )),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        // Setters mutate __time on args[0] and return the new ms.
        "setTime" => {
            let new_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            if let Some(Value::Object(obj)) = args.first() {
                obj.lock().unwrap().properties.insert("__time".into(), Value::F64(new_ms));
            }
            Value::F64(new_ms)
        }
        "setFullYear" | "setUTCFullYear" => date_setter(args, "year"),
        "setMonth" | "setUTCMonth" => date_setter(args, "month"),
        "setDate" | "setUTCDate" => date_setter(args, "day"),
        "setHours" | "setUTCHours" => date_setter(args, "hour"),
        "setMinutes" | "setUTCMinutes" => date_setter(args, "minute"),
        "setSeconds" | "setUTCSeconds" => date_setter(args, "second"),
        "setMilliseconds" | "setUTCMilliseconds" => date_setter(args, "millisecond"),
        _ => return None,
    };
    Some(result)
}

fn date_setter(args: &[Value], component: &str) -> Value {
    let new_ms = setter_helper(args, component);
    if let Some(Value::Object(obj)) = args.first() {
        obj.lock().unwrap().properties.insert("__time".into(), Value::F64(new_ms));
    }
    Value::F64(new_ms)
}

/// Flexible date-string parser. Accepts, in order:
/// 1. RFC-3339 / ISO-8601 full timestamps
/// 2. `YYYY-MM-DD HH:MM:SS`
/// 3. `YYYY-MM-DD`
/// 4. `YYYY/MM/DD` variants
/// 5. Literal `"now"`
/// Returns `ms` since Unix epoch, or `f64::NAN` on failure (ECMA
/// `Date.parse` spec).
fn parse_natural(s: &str) -> f64 {
    let s = s.trim();
    if s.eq_ignore_ascii_case("now") {
        return ms_of(Utc::now());
    }
    if let Ok(dt) = DateTime::parse_from_rfc3339(s) {
        return ms_of(dt.with_timezone(&Utc));
    }
    for pat in &[
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    ] {
        if let Ok(ndt) = NaiveDateTime::parse_from_str(s, pat) {
            return ms_of(Utc.from_utc_datetime(&ndt));
        }
    }
    for pat in &["%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d-%m-%Y"] {
        if let Ok(nd) = NaiveDate::parse_from_str(s, pat) {
            let ndt = nd.and_time(NaiveTime::from_hms_opt(0, 0, 0).unwrap());
            return ms_of(Utc.from_utc_datetime(&ndt));
        }
    }
    f64::NAN
}

pub fn register(vm: &mut VM) {
    // ── ECMA-262 §21.4 Date primitives ────────────────────────────────
    //
    // The underlying timestamp source is `wasi:clocks/wall-clock.now`
    // (WASI 0.2.11 — see proposals/WASI/proposals/clocks/wit/wall-clock.wit).
    // ECMA Date is the §21.4 adapter on top: it normalizes the WASI
    // `datetime` record `{ seconds, nanoseconds }` into ECMA's
    // [[DateValue]] millisecond representation. Calling
    // `wasi:clocks/wall-clock.now` directly through the host_registry
    // would require Promise-shaped indirection (host_fns are by index);
    // instead we share the underlying `SystemTime::now()` source — the
    // WASI host fn and the inline calls below produce identical
    // observable timestamps. Read the WASI fn for the spec record shape.

    // Date.now() → ms since epoch (ECMA §21.4.1.1).
    vm.register_host_fn(MODULE, "now", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::F64(ms_of(Utc::now()))
    }));

    // new Date(...) — ECMA §21.4.2 constructor.
    //   new Date()                  → wall-clock.now() lifted into a Date object
    //   new Date(ms)                → ms-keyed Date object
    //   new Date(year, month, ...)  → constructed via §21.4.2.2 component form
    //
    // Returns an Object stamped `__type=Date` with `__time` (ms since
    // epoch) — the same shape every other `ecma:date.*` method
    // operates on.
    vm.register_host_fn(MODULE, "new", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = match args.len() {
            // No args (or just `this` from `New Date()` paths): wall-clock.now()
            0 => ms_of(Utc::now()),
            1 => {
                // Single arg: either `this` from a New X path (Object) or
                // a numeric ms / parseable string per §21.4.2.1.
                match args.first() {
                    Some(Value::Object(_)) => ms_of(Utc::now()),
                    Some(Value::String(s)) => parse_natural(s.as_ref()),
                    Some(v) => v.as_f64(),
                    None => ms_of(Utc::now()),
                }
            }
            _ => {
                // Multi-arg (year, month, day?, ...): mirror Date.UTC arity.
                // Skip `this` if first arg is an Object (the .NET wrapper
                // path passes `this` as arg 0).
                let offset = if matches!(args.first(), Some(Value::Object(_))) { 1 } else { 0 };
                let y = args.get(offset).map(|v| v.as_f64() as i32).unwrap_or(1970);
                let m = args.get(offset + 1).map(|v| v.as_f64() as u32 + 1).unwrap_or(1);
                let d = args.get(offset + 2).map(|v| v.as_f64() as u32).unwrap_or(1);
                let h = args.get(offset + 3).map(|v| v.as_f64() as u32).unwrap_or(0);
                let mn = args.get(offset + 4).map(|v| v.as_f64() as u32).unwrap_or(0);
                let s = args.get(offset + 5).map(|v| v.as_f64() as u32).unwrap_or(0);
                let mss = args.get(offset + 6).map(|v| v.as_f64() as u32).unwrap_or(0);
                match (
                    NaiveDate::from_ymd_opt(y, m, d),
                    NaiveTime::from_hms_milli_opt(h, mn, s, mss),
                ) {
                    (Some(nd), Some(nt)) => ms_of(Utc.from_utc_datetime(&nd.and_time(nt))),
                    _ => f64::NAN,
                }
            }
        };
        let mut obj = vybe_bytecode::value::Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Date")));
        obj.properties.insert("__time".into(), Value::F64(ms));
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)))
    }));

    // Date.parse(str) → ms since epoch, NaN on failure.
    vm.register_host_fn(MODULE, "parse", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::F64(parse_natural(&s))
    }));

    // Date.UTC(year, month, day?, hours?, minutes?, seconds?, ms?) → ms
    // Spec: month is 0-indexed. Defaults: day=1, hours/min/sec/ms=0.
    vm.register_host_fn(MODULE, "UTC", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let y = args.first().map(|v| v.as_f64() as i32).unwrap_or(1970);
        let m = args.get(1).map(|v| v.as_f64() as u32 + 1).unwrap_or(1);
        let d = args.get(2).map(|v| v.as_f64() as u32).unwrap_or(1);
        let h = args.get(3).map(|v| v.as_f64() as u32).unwrap_or(0);
        let min = args.get(4).map(|v| v.as_f64() as u32).unwrap_or(0);
        let s = args.get(5).map(|v| v.as_f64() as u32).unwrap_or(0);
        let ms = args.get(6).map(|v| v.as_f64() as u32).unwrap_or(0);
        let nd = NaiveDate::from_ymd_opt(y, m, d);
        let nt = NaiveTime::from_hms_milli_opt(h, min, s, ms);
        match (nd, nt) {
            (Some(nd), Some(nt)) => {
                Value::F64(ms_of(Utc.from_utc_datetime(&nd.and_time(nt))))
            }
            _ => Value::F64(f64::NAN),
        }
    }));

    // getFullYear / getMonth / getDate / getDay / getHours / getMinutes
    // / getSeconds — operate on a ms timestamp. `getMonth` is 0-indexed
    // per ECMA, `getDay` is day-of-week (Sunday=0).
    macro_rules! getter {
        ($name:literal, $body:expr) => {
            vm.register_host_fn(MODULE, $name, Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                let ms = ms_arg(args, 0);
                if let Some(dt) = dt_from_ms(ms) {
                    Value::F64($body(&dt) as f64)
                } else {
                    Value::F64(f64::NAN)
                }
            }));
        };
    }
    getter!("getFullYear", |dt: &DateTime<Utc>| dt.year());
    getter!("getMonth", |dt: &DateTime<Utc>| dt.month() as i32 - 1);
    getter!("getDate", |dt: &DateTime<Utc>| dt.day() as i32);
    getter!("getDay", |dt: &DateTime<Utc>| dt.weekday().num_days_from_sunday() as i32);
    getter!("getHours", |dt: &DateTime<Utc>| dt.hour() as i32);
    getter!("getMinutes", |dt: &DateTime<Utc>| dt.minute() as i32);
    getter!("getSeconds", |dt: &DateTime<Utc>| dt.second() as i32);
    getter!("getMilliseconds", |dt: &DateTime<Utc>| dt.timestamp_subsec_millis() as i32);
    getter!("getTime", |dt: &DateTime<Utc>| dt.timestamp_millis());
    // UTC variants — chrono is already in UTC, so they share the impl.
    getter!("getUTCFullYear", |dt: &DateTime<Utc>| dt.year());
    getter!("getUTCMonth", |dt: &DateTime<Utc>| dt.month() as i32 - 1);
    getter!("getUTCDate", |dt: &DateTime<Utc>| dt.day() as i32);
    getter!("getUTCDay", |dt: &DateTime<Utc>| dt.weekday().num_days_from_sunday() as i32);
    getter!("getUTCHours", |dt: &DateTime<Utc>| dt.hour() as i32);
    getter!("getUTCMinutes", |dt: &DateTime<Utc>| dt.minute() as i32);
    getter!("getUTCSeconds", |dt: &DateTime<Utc>| dt.second() as i32);
    getter!("getUTCMilliseconds", |dt: &DateTime<Utc>| dt.timestamp_subsec_millis() as i32);
    getter!("valueOf", |dt: &DateTime<Utc>| dt.timestamp_millis());
    getter!("getTimezoneOffset", |_dt: &DateTime<Utc>| 0i32);

    // toISOString(this) — "2026-03-25T00:00:00.000Z"
    vm.register_host_fn(MODULE, "toISOString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = ms_arg(args, 0);
        match dt_from_ms(ms) {
            Some(dt) => {
                let s = dt.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string();
                Value::String(Arc::from(s.as_str()))
            }
            None => Value::String(Arc::from("Invalid Date")),
        }
    }));

    // toString(this) — "Mon Mar 25 2026 00:00:00 GMT+0000"
    vm.register_host_fn(MODULE, "toString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = ms_arg(args, 0);
        match dt_from_ms(ms) {
            Some(dt) => {
                let s = dt.format("%a %b %d %Y %H:%M:%S GMT+0000 (UTC)").to_string();
                Value::String(Arc::from(s.as_str()))
            }
            None => Value::String(Arc::from("Invalid Date")),
        }
    }));

    // toDateString(this) — "Mon Mar 25 2026" (date portion only)
    vm.register_host_fn(MODULE, "toDateString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = ms_arg(args, 0);
        match dt_from_ms(ms) {
            Some(dt) => Value::String(Arc::from(dt.format("%a %b %d %Y").to_string().as_str())),
            None => Value::String(Arc::from("Invalid Date")),
        }
    }));

    // toTimeString(this) — "00:00:00 GMT+0000 (UTC)"
    vm.register_host_fn(MODULE, "toTimeString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = ms_arg(args, 0);
        match dt_from_ms(ms) {
            Some(dt) => Value::String(Arc::from(dt.format("%H:%M:%S GMT+0000 (UTC)").to_string().as_str())),
            None => Value::String(Arc::from("Invalid Date")),
        }
    }));

    // toJSON(this) — same as toISOString per ECMA-262 §21.4.4.37
    vm.register_host_fn(MODULE, "toJSON", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = ms_arg(args, 0);
        match dt_from_ms(ms) {
            Some(dt) => Value::String(Arc::from(dt.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string().as_str())),
            None => Value::Null,
        }
    }));

    // setTime(this, ms) — mutates `__time` and returns the new ms.
    vm.register_host_fn(MODULE, "setTime", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let new_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(f64::NAN);
        if let Some(Value::Object(obj)) = args.first() {
            obj.lock().unwrap().properties.insert("__time".into(), Value::F64(new_ms));
        }
        Value::F64(new_ms)
    }));

    // Component setters — mutate the Date's __time and return new ms.
    macro_rules! setter {
        ($name:literal, $component:ident) => {
            vm.register_host_fn(MODULE, $name, Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                let new_ms = setter_helper(args, stringify!($component));
                if let Some(Value::Object(obj)) = args.first() {
                    obj.lock().unwrap().properties.insert("__time".into(), Value::F64(new_ms));
                }
                Value::F64(new_ms)
            }));
        };
    }
    setter!("setFullYear", year);
    setter!("setMonth", month);
    setter!("setDate", day);
    setter!("setHours", hour);
    setter!("setMinutes", minute);
    setter!("setSeconds", second);
    setter!("setMilliseconds", millisecond);
    setter!("setUTCFullYear", year);
    setter!("setUTCMonth", month);
    setter!("setUTCDate", day);
    setter!("setUTCHours", hour);
    setter!("setUTCMinutes", minute);
    setter!("setUTCSeconds", second);
    setter!("setUTCMilliseconds", millisecond);

    // ── Cross-language adapters (Vybe extension, NOT ECMA) ─────────────

    // fromUnixSeconds(secs) → ms. Bridges POSIX `time()` to JS Date
    // model. Fractional seconds preserved.
    vm.register_host_fn(MODULE, "fromUnixSeconds", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let secs = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
        Value::F64(secs * 1000.0)
    }));

    // toUnixSeconds(ms) → floor(ms / 1000). Bridges JS Date → POSIX.
    vm.register_host_fn(MODULE, "toUnixSeconds", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64()).unwrap_or_else(|| ms_of(Utc::now()));
        Value::F64((ms / 1000.0).floor())
    }));

    // nowSeconds() — POSIX-style: current seconds since epoch. Equivalent
    // to `toUnixSeconds(now())` but saves a call.
    vm.register_host_fn(MODULE, "nowSeconds", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::F64(Utc::now().timestamp() as f64)
    }));

}
