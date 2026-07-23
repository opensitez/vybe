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

use chrono::{
    DateTime, Datelike, Duration, NaiveDate, NaiveDateTime, NaiveTime, TimeZone, Timelike, Utc,
};
use std::sync::{Arc, Mutex, OnceLock};
use vybe_bytecode::value::Object;
use vybe_bytecode::{HostContext, Value, VM};

const MODULE: &str = "ecma:date";

static DATE_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

/// Canonical `%Date.prototype%`. A singleton so the global wiring (which
/// populates the methods) and the `new Date()` constructor (which links
/// each instance's `__proto__`) reference the SAME object — making
/// `Object.getPrototypeOf(new Date()) === Date.prototype` hold.
pub(crate) fn shared_date_prototype() -> Value {
    Value::Object(
        DATE_PROTOTYPE
            .get_or_init(|| Arc::new(Mutex::new(Object::new())))
            .clone(),
    )
}

fn dt_from_ms(ms: f64) -> Option<DateTime<Utc>> {
    if !ms.is_finite() {
        return None;
    }
    let secs = (ms / 1000.0).floor() as i64;
    let nsecs = (((ms.rem_euclid(1000.0)) * 1_000_000.0) as u32).min(999_999_999);
    Utc.timestamp_opt(secs, nsecs).single()
}

fn ms_of(dt: DateTime<Utc>) -> f64 {
    (dt.timestamp() as f64) * 1000.0 + (dt.timestamp_subsec_millis() as f64)
}

fn format_utc_string(ms: f64) -> Option<String> {
    dt_from_ms(ms).map(|dt| dt.format("%a, %d %b %Y %H:%M:%S GMT").to_string())
}

fn component_i64(args: &[Value], idx: usize) -> Result<Option<i64>, ()> {
    match args.get(idx) {
        Some(value) => {
            let numeric = value.as_f64();
            if numeric.is_nan() {
                Err(())
            } else {
                Ok(Some(numeric.trunc() as i64))
            }
        }
        None => Ok(None),
    }
}

fn build_utc_ms(
    year: i32,
    month_zero: i64,
    day: i64,
    hour: i64,
    minute: i64,
    second: i64,
    millisecond: i64,
) -> f64 {
    let total_months = (year as i64) * 12 + month_zero;
    let normalized_year = total_months.div_euclid(12) as i32;
    let normalized_month = (total_months.rem_euclid(12) + 1) as u32;
    let Some(month_start) = NaiveDate::from_ymd_opt(normalized_year, normalized_month, 1) else {
        return f64::NAN;
    };
    let Some(midnight) = month_start.and_hms_milli_opt(0, 0, 0, 0) else {
        return f64::NAN;
    };
    let dt = Utc.from_utc_datetime(&midnight)
        + Duration::days(day - 1)
        + Duration::hours(hour)
        + Duration::minutes(minute)
        + Duration::seconds(second)
        + Duration::milliseconds(millisecond);
    ms_of(dt)
}

fn construct_date_from_args(values: &[Value]) -> f64 {
    let year = values.first().map(|v| v.as_f64() as i32).unwrap_or(1970);
    let constructor_year = if (0..=99).contains(&year) {
        year + 1900
    } else {
        year
    };
    let month_zero = values
        .get(1)
        .map(|v| v.as_f64().trunc() as i64)
        .unwrap_or(0);
    let day = values
        .get(2)
        .map(|v| v.as_f64().trunc() as i64)
        .unwrap_or(1);
    let hour = values
        .get(3)
        .map(|v| v.as_f64().trunc() as i64)
        .unwrap_or(0);
    let minute = values
        .get(4)
        .map(|v| v.as_f64().trunc() as i64)
        .unwrap_or(0);
    let second = values
        .get(5)
        .map(|v| v.as_f64().trunc() as i64)
        .unwrap_or(0);
    let millisecond = values
        .get(6)
        .map(|v| v.as_f64().trunc() as i64)
        .unwrap_or(0);
    build_utc_ms(
        constructor_year,
        month_zero,
        day,
        hour,
        minute,
        second,
        millisecond,
    )
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
    let dt = match dt_from_ms(ms) {
        Some(d) => d,
        None => return f64::NAN,
    };
    let mut year = dt.year();
    let mut month_zero = dt.month0() as i64;
    let mut day = dt.day() as i64;
    let mut hour = dt.hour() as i64;
    let mut minute = dt.minute() as i64;
    let mut second = dt.second() as i64;
    let mut millisecond = dt.timestamp_subsec_millis() as i64;

    match component {
        "year" => {
            let Some(next_year) = component_i64(args, 1).ok().flatten() else {
                return f64::NAN;
            };
            year = next_year as i32;
            if let Ok(Some(next_month)) = component_i64(args, 2) {
                month_zero = next_month;
            } else if matches!(component_i64(args, 2), Err(())) {
                return f64::NAN;
            }
            if let Ok(Some(next_day)) = component_i64(args, 3) {
                day = next_day;
            } else if matches!(component_i64(args, 3), Err(())) {
                return f64::NAN;
            }
        }
        "month" => {
            let Some(next_month) = component_i64(args, 1).ok().flatten() else {
                return f64::NAN;
            };
            month_zero = next_month;
            if let Ok(Some(next_day)) = component_i64(args, 2) {
                day = next_day;
            } else if matches!(component_i64(args, 2), Err(())) {
                return f64::NAN;
            }
        }
        "day" => {
            let Some(next_day) = component_i64(args, 1).ok().flatten() else {
                return f64::NAN;
            };
            day = next_day;
        }
        "hour" => {
            let Some(next_hour) = component_i64(args, 1).ok().flatten() else {
                return f64::NAN;
            };
            hour = next_hour;
            if let Ok(Some(next_minute)) = component_i64(args, 2) {
                minute = next_minute;
            } else if matches!(component_i64(args, 2), Err(())) {
                return f64::NAN;
            }
            if let Ok(Some(next_second)) = component_i64(args, 3) {
                second = next_second;
            } else if matches!(component_i64(args, 3), Err(())) {
                return f64::NAN;
            }
            if let Ok(Some(next_millisecond)) = component_i64(args, 4) {
                millisecond = next_millisecond;
            } else if matches!(component_i64(args, 4), Err(())) {
                return f64::NAN;
            }
        }
        "minute" => {
            let Some(next_minute) = component_i64(args, 1).ok().flatten() else {
                return f64::NAN;
            };
            minute = next_minute;
            if let Ok(Some(next_second)) = component_i64(args, 2) {
                second = next_second;
            } else if matches!(component_i64(args, 2), Err(())) {
                return f64::NAN;
            }
            if let Ok(Some(next_millisecond)) = component_i64(args, 3) {
                millisecond = next_millisecond;
            } else if matches!(component_i64(args, 3), Err(())) {
                return f64::NAN;
            }
        }
        "second" => {
            let Some(next_second) = component_i64(args, 1).ok().flatten() else {
                return f64::NAN;
            };
            second = next_second;
            if let Ok(Some(next_millisecond)) = component_i64(args, 2) {
                millisecond = next_millisecond;
            } else if matches!(component_i64(args, 2), Err(())) {
                return f64::NAN;
            }
        }
        "millisecond" => {
            let Some(next_millisecond) = component_i64(args, 1).ok().flatten() else {
                return f64::NAN;
            };
            millisecond = next_millisecond;
        }
        _ => return f64::NAN,
    }

    build_utc_ms(year, month_zero, day, hour, minute, second, millisecond)
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
        "getYear" => getter!(|dt: &DateTime<Utc>| dt.year() - 1900),
        "getMonth" | "getUTCMonth" => getter!(|dt: &DateTime<Utc>| dt.month() as i32 - 1),
        "getDate" | "getUTCDate" => getter!(|dt: &DateTime<Utc>| dt.day() as i32),
        "getDay" | "getUTCDay" => {
            getter!(|dt: &DateTime<Utc>| dt.weekday().num_days_from_sunday() as i32)
        }
        "getHours" | "getUTCHours" => getter!(|dt: &DateTime<Utc>| dt.hour() as i32),
        "getMinutes" | "getUTCMinutes" => getter!(|dt: &DateTime<Utc>| dt.minute() as i32),
        "getSeconds" | "getUTCSeconds" => getter!(|dt: &DateTime<Utc>| dt.second() as i32),
        "getMilliseconds" | "getUTCMilliseconds" => {
            getter!(|dt: &DateTime<Utc>| dt.timestamp_subsec_millis() as i32)
        }
        "getTime" | "valueOf" => Value::F64(ms_arg(args, 0)),
        "getTimezoneOffset" => Value::F64(0.0),
        "toISOString" | "toJSON" => {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string().as_str(),
                )),
                None if method == "toJSON" => Value::Null,
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        "toUTCString" => {
            let ms = ms_arg(args, 0);
            match format_utc_string(ms) {
                Some(text) => Value::String(Arc::from(text.as_str())),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        "toString" | "toLocaleString" => {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%a %b %d %Y %H:%M:%S GMT+0000 (UTC)")
                        .to_string()
                        .as_str(),
                )),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        "toDateString" | "toLocaleDateString" => {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(dt.format("%a %b %d %Y").to_string().as_str())),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        "toTimeString" | "toLocaleTimeString" => {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%H:%M:%S GMT+0000 (UTC)").to_string().as_str(),
                )),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }
        // Setters mutate __time on args[0] and return the new ms.
        "setTime" => {
            let new_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            if let Some(Value::Object(obj)) = args.first() {
                obj.lock()
                    .unwrap()
                    .properties
                    .insert("__time".into(), Value::F64(new_ms));
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
        obj.lock()
            .unwrap()
            .properties
            .insert("__time".into(), Value::F64(new_ms));
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
    if let Ok(dt) = DateTime::parse_from_rfc2822(s) {
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
    for pat in &[
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%d-%m-%Y",
        "%b %e, %Y",
        "%B %e, %Y",
    ] {
        if let Ok(nd) = NaiveDate::parse_from_str(s, pat) {
            let ndt = nd.and_time(NaiveTime::from_hms_opt(0, 0, 0).unwrap());
            return ms_of(Utc.from_utc_datetime(&ndt));
        }
    }
    if s.chars().all(|ch| ch.is_ascii_digit()) {
        if let Ok(year) = s.parse::<i32>() {
            return build_utc_ms(year, 0, 1, 0, 0, 0, 0);
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
    vm.register_host_fn(
        MODULE,
        "now",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(ms_of(Utc::now()))),
    );

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
                // Single arg: either `this` from a New X path (plain Object),
                // a spread array from `new Date(...args)` (Array Object),
                // a numeric ms, or a parseable string per §21.4.2.1.
                match args.first() {
                    Some(Value::Object(obj)) => {
                        let o = obj.lock().unwrap();
                        if let vybe_bytecode::value::ObjectKind::Array(elems) = &o.kind {
                            construct_date_from_args(elems)
                        } else if matches!(o.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == "Date") {
                            o.properties.get("__time").map(Value::as_f64).unwrap_or(f64::NAN)
                        } else {
                            // Plain Object `this` from .NET/VB New Date() path
                            ms_of(Utc::now())
                        }
                    }
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
                construct_date_from_args(&args[offset..])
            }
        };
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("Date")));
        obj.properties.insert("__time".into(), Value::F64(ms));
        obj.properties
            .insert("__proto__".into(), shared_date_prototype());
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // Date.parse(str) → ms since epoch, NaN on failure.
    vm.register_host_fn(
        MODULE,
        "parse",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            Value::F64(parse_natural(&s))
        }),
    );

    // Date.UTC(year, month, day?, hours?, minutes?, seconds?, ms?) → ms
    // Spec: month is 0-indexed. Defaults: day=1, hours/min/sec/ms=0.
    vm.register_host_fn(
        MODULE,
        "UTC",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::F64(construct_date_from_args(args))
        }),
    );

    // getFullYear / getMonth / getDate / getDay / getHours / getMinutes
    // / getSeconds — operate on a ms timestamp. `getMonth` is 0-indexed
    // per ECMA, `getDay` is day-of-week (Sunday=0).
    macro_rules! getter {
        ($name:literal, $body:expr) => {
            vm.register_host_fn(
                MODULE,
                $name,
                Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                    let ms = ms_arg(args, 0);
                    if let Some(dt) = dt_from_ms(ms) {
                        Value::F64($body(&dt) as f64)
                    } else {
                        Value::F64(f64::NAN)
                    }
                }),
            );
        };
    }
    getter!("getFullYear", |dt: &DateTime<Utc>| dt.year());
    getter!("getYear", |dt: &DateTime<Utc>| dt.year() - 1900);
    getter!("getMonth", |dt: &DateTime<Utc>| dt.month() as i32 - 1);
    getter!("getDate", |dt: &DateTime<Utc>| dt.day() as i32);
    getter!(
        "getDay",
        |dt: &DateTime<Utc>| dt.weekday().num_days_from_sunday() as i32
    );
    getter!("getHours", |dt: &DateTime<Utc>| dt.hour() as i32);
    getter!("getMinutes", |dt: &DateTime<Utc>| dt.minute() as i32);
    getter!("getSeconds", |dt: &DateTime<Utc>| dt.second() as i32);
    getter!(
        "getMilliseconds",
        |dt: &DateTime<Utc>| dt.timestamp_subsec_millis() as i32
    );
    getter!("getTime", |dt: &DateTime<Utc>| dt.timestamp_millis());
    // UTC variants — chrono is already in UTC, so they share the impl.
    getter!("getUTCFullYear", |dt: &DateTime<Utc>| dt.year());
    getter!("getUTCMonth", |dt: &DateTime<Utc>| dt.month() as i32 - 1);
    getter!("getUTCDate", |dt: &DateTime<Utc>| dt.day() as i32);
    getter!(
        "getUTCDay",
        |dt: &DateTime<Utc>| dt.weekday().num_days_from_sunday() as i32
    );
    getter!("getUTCHours", |dt: &DateTime<Utc>| dt.hour() as i32);
    getter!("getUTCMinutes", |dt: &DateTime<Utc>| dt.minute() as i32);
    getter!("getUTCSeconds", |dt: &DateTime<Utc>| dt.second() as i32);
    getter!(
        "getUTCMilliseconds",
        |dt: &DateTime<Utc>| dt.timestamp_subsec_millis() as i32
    );
    getter!("valueOf", |dt: &DateTime<Utc>| dt.timestamp_millis());
    getter!("getTimezoneOffset", |_dt: &DateTime<Utc>| 0i32);

    // toISOString(this) — "2026-03-25T00:00:00.000Z"
    vm.register_host_fn(
        MODULE,
        "toISOString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => {
                    let s = dt.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string();
                    Value::String(Arc::from(s.as_str()))
                }
                None => Value::String(Arc::from("Invalid Date")),
            }
        }),
    );

    // toString(this) — "Mon Mar 25 2026 00:00:00 GMT+0000"
    vm.register_host_fn(
        MODULE,
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => {
                    let s = dt.format("%a %b %d %Y %H:%M:%S GMT+0000 (UTC)").to_string();
                    Value::String(Arc::from(s.as_str()))
                }
                None => Value::String(Arc::from("Invalid Date")),
            }
        }),
    );

    vm.register_host_fn(
        MODULE,
        "toUTCString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ms = ms_arg(args, 0);
            match format_utc_string(ms) {
                Some(text) => Value::String(Arc::from(text.as_str())),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }),
    );

    // toDateString(this) — "Mon Mar 25 2026" (date portion only)
    vm.register_host_fn(
        MODULE,
        "toDateString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(dt.format("%a %b %d %Y").to_string().as_str())),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }),
    );

    // toTimeString(this) — "00:00:00 GMT+0000 (UTC)"
    vm.register_host_fn(
        MODULE,
        "toTimeString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%H:%M:%S GMT+0000 (UTC)").to_string().as_str(),
                )),
                None => Value::String(Arc::from("Invalid Date")),
            }
        }),
    );

    // toJSON(this) — same as toISOString per ECMA-262 §21.4.4.37
    vm.register_host_fn(
        MODULE,
        "toJSON",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ms = ms_arg(args, 0);
            match dt_from_ms(ms) {
                Some(dt) => Value::String(Arc::from(
                    dt.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string().as_str(),
                )),
                None => Value::Null,
            }
        }),
    );

    // setTime(this, ms) — mutates `__time` and returns the new ms.
    vm.register_host_fn(
        MODULE,
        "setTime",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let new_ms = args.get(1).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            if let Some(Value::Object(obj)) = args.first() {
                obj.lock()
                    .unwrap()
                    .properties
                    .insert("__time".into(), Value::F64(new_ms));
            }
            Value::F64(new_ms)
        }),
    );

    // Component setters — mutate the Date's __time and return new ms.
    macro_rules! setter {
        ($name:literal, $component:ident) => {
            vm.register_host_fn(
                MODULE,
                $name,
                Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                    let new_ms = setter_helper(args, stringify!($component));
                    if let Some(Value::Object(obj)) = args.first() {
                        obj.lock()
                            .unwrap()
                            .properties
                            .insert("__time".into(), Value::F64(new_ms));
                    }
                    Value::F64(new_ms)
                }),
            );
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
    vm.register_host_fn(
        MODULE,
        "fromUnixSeconds",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let secs = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            Value::F64(secs * 1000.0)
        }),
    );

    // toUnixSeconds(ms) → floor(ms / 1000). Bridges JS Date → POSIX.
    vm.register_host_fn(
        MODULE,
        "toUnixSeconds",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let ms = args
                .first()
                .map(|v| v.as_f64())
                .unwrap_or_else(|| ms_of(Utc::now()));
            Value::F64((ms / 1000.0).floor())
        }),
    );

    // nowSeconds() — POSIX-style: current seconds since epoch. Equivalent
    // to `toUnixSeconds(now())` but saves a call.
    vm.register_host_fn(
        MODULE,
        "nowSeconds",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::F64(Utc::now().timestamp() as f64)
        }),
    );

    vm.register_host_fn(
        MODULE,
        "toLocaleString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            dispatch_date_method("toString", args)
                .unwrap_or_else(|| Value::String(Arc::from("Invalid Date")))
        }),
    );

    vm.register_host_fn(
        MODULE,
        "toLocaleDateString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            dispatch_date_method("toDateString", args)
                .unwrap_or_else(|| Value::String(Arc::from("Invalid Date")))
        }),
    );

    vm.register_host_fn(
        MODULE,
        "toLocaleTimeString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            dispatch_date_method("toTimeString", args)
                .unwrap_or_else(|| Value::String(Arc::from("Invalid Date")))
        }),
    );
}
