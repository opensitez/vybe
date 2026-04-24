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

/// Translate a PHP `date()` format string to output against `dt`.
/// Supports the common subset — Y y m n d j H G i s U l D F M N w t L a A.
fn format_php(fmt: &str, dt: &DateTime<Utc>) -> String {
    let mut out = String::with_capacity(fmt.len() * 2);
    let mut chars = fmt.chars().peekable();
    while let Some(c) = chars.next() {
        // Backslash escapes the next character literally (PHP convention).
        if c == '\\' {
            if let Some(next) = chars.next() {
                out.push(next);
            }
            continue;
        }
        match c {
            'Y' => out.push_str(&format!("{:04}", dt.year())),
            'y' => out.push_str(&format!("{:02}", dt.year() % 100)),
            'm' => out.push_str(&format!("{:02}", dt.month())),
            'n' => out.push_str(&format!("{}", dt.month())),
            'd' => out.push_str(&format!("{:02}", dt.day())),
            'j' => out.push_str(&format!("{}", dt.day())),
            'H' => out.push_str(&format!("{:02}", dt.hour())),
            'G' => out.push_str(&format!("{}", dt.hour())),
            'h' => {
                let h12 = ((dt.hour() % 12) as i64).max(1);
                let h = if dt.hour() % 12 == 0 { 12 } else { h12 as u32 };
                out.push_str(&format!("{:02}", h));
            }
            'g' => {
                let h = if dt.hour() % 12 == 0 { 12 } else { dt.hour() % 12 };
                out.push_str(&format!("{}", h));
            }
            'i' => out.push_str(&format!("{:02}", dt.minute())),
            's' => out.push_str(&format!("{:02}", dt.second())),
            'a' => out.push_str(if dt.hour() < 12 { "am" } else { "pm" }),
            'A' => out.push_str(if dt.hour() < 12 { "AM" } else { "PM" }),
            'U' => out.push_str(&format!("{}", dt.timestamp())),
            'l' => out.push_str(match dt.weekday().num_days_from_sunday() {
                0 => "Sunday", 1 => "Monday", 2 => "Tuesday",
                3 => "Wednesday", 4 => "Thursday", 5 => "Friday", _ => "Saturday",
            }),
            'D' => out.push_str(match dt.weekday().num_days_from_sunday() {
                0 => "Sun", 1 => "Mon", 2 => "Tue",
                3 => "Wed", 4 => "Thu", 5 => "Fri", _ => "Sat",
            }),
            'F' => out.push_str(match dt.month() {
                1 => "January", 2 => "February", 3 => "March",
                4 => "April", 5 => "May", 6 => "June",
                7 => "July", 8 => "August", 9 => "September",
                10 => "October", 11 => "November", _ => "December",
            }),
            'M' => out.push_str(match dt.month() {
                1 => "Jan", 2 => "Feb", 3 => "Mar",
                4 => "Apr", 5 => "May", 6 => "Jun",
                7 => "Jul", 8 => "Aug", 9 => "Sep",
                10 => "Oct", 11 => "Nov", _ => "Dec",
            }),
            'N' => {
                // ISO-8601 numeric day: Monday=1 … Sunday=7.
                let iso = dt.weekday().number_from_monday();
                out.push_str(&format!("{}", iso));
            }
            'w' => out.push_str(&format!("{}", dt.weekday().num_days_from_sunday())),
            't' => {
                let (y, m) = (dt.year(), dt.month());
                let next_month = if m == 12 {
                    NaiveDate::from_ymd_opt(y + 1, 1, 1)
                } else {
                    NaiveDate::from_ymd_opt(y, m + 1, 1)
                };
                let first = NaiveDate::from_ymd_opt(y, m, 1);
                let days = match (first, next_month) {
                    (Some(f), Some(n)) => (n.signed_duration_since(f).num_days()) as i64,
                    _ => 30,
                };
                out.push_str(&format!("{}", days));
            }
            'L' => {
                let y = dt.year();
                let leap = y % 4 == 0 && (y % 100 != 0 || y % 400 == 0);
                out.push_str(if leap { "1" } else { "0" });
            }
            'z' => {
                let doy = dt.ordinal0();
                out.push_str(&format!("{}", doy));
            }
            'T' => out.push_str("UTC"),
            'c' => out.push_str(&dt.to_rfc3339()),
            'r' => out.push_str(&dt.to_rfc2822()),
            other => out.push(other),
        }
    }
    out
}

/// C `strftime` format — `%Y`, `%m`, etc. Delegates to chrono's format
/// which implements the C spec faithfully.
fn format_strftime(fmt: &str, dt: &DateTime<Utc>) -> String {
    dt.format(fmt).to_string()
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

    // Date.now() → ms since epoch.
    vm.register_host_fn(MODULE, "now", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::F64(ms_of(Utc::now()))
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
                let ms = args.first().map(|v| v.as_f64()).unwrap_or_else(|| ms_of(Utc::now()));
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
    getter!("getTime", |dt: &DateTime<Utc>| dt.timestamp_millis() as i64 as i32);

    // toISOString(ms) — "2026-03-25T00:00:00.000Z"
    vm.register_host_fn(MODULE, "toISOString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64()).unwrap_or_else(|| ms_of(Utc::now()));
        match dt_from_ms(ms) {
            Some(dt) => {
                let s = dt.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string();
                Value::String(Arc::from(s.as_str()))
            }
            None => Value::String(Arc::from("Invalid Date")),
        }
    }));

    // toString(ms) — "Mon Mar 25 2026 00:00:00 GMT+0000"
    vm.register_host_fn(MODULE, "toString", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64()).unwrap_or_else(|| ms_of(Utc::now()));
        match dt_from_ms(ms) {
            Some(dt) => {
                let s = dt.format("%a %b %d %Y %H:%M:%S GMT+0000 (UTC)").to_string();
                Value::String(Arc::from(s.as_str()))
            }
            None => Value::String(Arc::from("Invalid Date")),
        }
    }));

    // ── Cross-language adapters (Vybe extension, NOT ECMA) ─────────────

    // formatPhp(ms, php_format_string) — PHP date()-style. Takes ms
    // epoch so PHP callers must first upscale from their seconds-based
    // time() via `fromUnixSeconds`.
    vm.register_host_fn(MODULE, "formatPhp", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64()).unwrap_or_else(|| ms_of(Utc::now()));
        let fmt = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        match dt_from_ms(ms) {
            Some(dt) => Value::String(Arc::from(format_php(&fmt, &dt).as_str())),
            None => Value::String(Arc::from("")),
        }
    }));

    // formatStrftime(ms, strftime_format) — C strftime %-codes.
    vm.register_host_fn(MODULE, "formatStrftime", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let ms = args.first().map(|v| v.as_f64()).unwrap_or_else(|| ms_of(Utc::now()));
        let fmt = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        match dt_from_ms(ms) {
            Some(dt) => Value::String(Arc::from(format_strftime(&fmt, &dt).as_str())),
            None => Value::String(Arc::from("")),
        }
    }));

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

    // ── Language-specific wrappers on seconds-epoch inputs ─────────────
    // These take a POSIX timestamp in seconds and a format string, same
    // signature as PHP's `date($fmt, $ts)` / `strftime($fmt, $ts)`.

    // phpDate(format, ts_secs) — PHP `date()` on seconds epoch.
    vm.register_host_fn(MODULE, "phpDate", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let fmt = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let ts = args.get(1)
            .map(|v| v.as_f64())
            .unwrap_or_else(|| Utc::now().timestamp() as f64);
        let dt = Utc.timestamp_opt(ts as i64, 0).single();
        match dt {
            Some(dt) => Value::String(Arc::from(format_php(&fmt, &dt).as_str())),
            None => Value::String(Arc::from("")),
        }
    }));

    // phpStrftime(format, ts_secs) — PHP `strftime()`.
    vm.register_host_fn(MODULE, "phpStrftime", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let fmt = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let ts = args.get(1)
            .map(|v| v.as_f64())
            .unwrap_or_else(|| Utc::now().timestamp() as f64);
        let dt = Utc.timestamp_opt(ts as i64, 0).single();
        match dt {
            Some(dt) => Value::String(Arc::from(format_strftime(&fmt, &dt).as_str())),
            None => Value::String(Arc::from("")),
        }
    }));

    // phpStrtotime(str) — PHP `strtotime()` → seconds epoch, or false-as-0
    // on parse failure (PHP returns `false`; we return 0 so arithmetic
    // callers don't need to null-check).
    vm.register_host_fn(MODULE, "phpStrtotime", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let ms = parse_natural(&s);
        if ms.is_nan() {
            Value::F64(0.0)
        } else {
            Value::F64((ms / 1000.0).floor())
        }
    }));

    // phpMktime(hour, minute, second, month, day, year) → seconds epoch.
    // PHP arg order is unusual — NOT (year, month, day, h, m, s).
    vm.register_host_fn(MODULE, "phpMktime", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let h = args.first().map(|v| v.as_f64() as u32).unwrap_or(0);
        let min = args.get(1).map(|v| v.as_f64() as u32).unwrap_or(0);
        let s = args.get(2).map(|v| v.as_f64() as u32).unwrap_or(0);
        let month = args.get(3).map(|v| v.as_f64() as u32).unwrap_or(1);
        let day = args.get(4).map(|v| v.as_f64() as u32).unwrap_or(1);
        let year = args.get(5).map(|v| v.as_f64() as i32).unwrap_or(1970);
        let nd = NaiveDate::from_ymd_opt(year, month, day);
        let nt = NaiveTime::from_hms_opt(h, min, s);
        match (nd, nt) {
            (Some(nd), Some(nt)) => {
                let ts = Utc.from_utc_datetime(&nd.and_time(nt)).timestamp();
                Value::F64(ts as f64)
            }
            _ => Value::F64(0.0),
        }
    }));
}
