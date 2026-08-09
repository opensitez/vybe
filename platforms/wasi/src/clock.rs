use std::sync::{Arc, OnceLock};
use std::time::Instant;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

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
            Value::Object(vybe_runtime::heap::alloc(obj))
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
            Value::Object(vybe_runtime::heap::alloc(obj))
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
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| system_clock_now()),
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
            Value::Object(vybe_runtime::heap::alloc(rec))
        }),
    );

    // ── wasi:clocks/system-clock — WASI 0.3.0 ───────────────────────────
    // `proposals/clocks/wit/system-clock.wit`. 0.3 renamed `wall-clock` to
    // `system-clock` and `resolution` to `get-resolution`; `now` still answers
    // a `{ seconds, nanoseconds }` record, now called `instant` rather than
    // `datetime`. Both spellings stay bound so either revision resolves.
    vm.register_host_fn(
        "wasi:clocks/system-clock",
        "now",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| system_clock_now()),
    );

    // `get-resolution: func() -> duration`, and `duration = u64` NANOSECONDS
    // (`clocks/wit/types.wit`) — a bare number, not the record 0.2's
    // `wall-clock.resolution` returned.
    vm.register_host_fn(
        "wasi:clocks/system-clock",
        "get-resolution",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(1.0)),
    );

    // ── wasi:clocks/timezone ────────────────────────────────────────────
    // display(when: datetime) → timezone-display { utc-offset: s32, name: string, in-daylight-saving-time: bool }
    // 0.2 only; 0.3 replaced it with `iana-id` + `to-debug-string`.
    vm.register_host_fn(
        "wasi:clocks/timezone",
        "display",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let seconds = instant_seconds(args.first());
            let offset = local_offset_seconds(seconds).unwrap_or(0);
            let mut rec = Object::new();
            rec.properties
                .insert("utc-offset".into(), Value::I32(offset as i32));
            rec.properties.insert(
                "name".into(),
                Value::String(Arc::from(
                    configured_zone_id()
                        .unwrap_or_else(|| "UTC".to_string())
                        .as_str(),
                )),
            );
            rec.properties.insert(
                "in-daylight-saving-time".into(),
                Value::Bool(local_is_dst(seconds).unwrap_or(false)),
            );
            Value::Object(vybe_runtime::heap::alloc(rec))
        }),
    );

    // `iana-id: func() -> option<string>` — the IANA Time Zone Database
    // identifier of the configured zone, or nothing when the host does not
    // expose one (`proposals/clocks/wit/timezone.wit`).
    vm.register_host_fn(
        "wasi:clocks/timezone",
        "iana-id",
        Box::new(
            |_ctx: &mut HostContext, _args: &[Value]| match configured_zone_id() {
                Some(id) => Value::String(Arc::from(id.as_str())),
                None => Value::Null,
            },
        ),
    );

    // `utc-offset: func(when: instant) -> option<s64>`.
    //
    // 0.3 changed BOTH the unit and the type: 0.2's was `s32` SECONDS and
    // always produced a value; 0.3's is `option<s64>` NANOSECONDS, and the
    // interface requires nothing back when the zone cannot be determined.
    // Returning a flat `0` would claim UTC for every host, so an unknown zone
    // answers null rather than a plausible lie.
    vm.register_host_fn(
        "wasi:clocks/timezone",
        "utc-offset",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let seconds = instant_seconds(args.first());
            match local_offset_seconds(seconds) {
                Some(offset) => Value::F64(offset as f64 * 1_000_000_000.0),
                None => Value::Null,
            }
        }),
    );

    // `to-debug-string: func() -> string` — for humans only; the spec warns
    // this must not be parsed. The IANA id when there is one, else the current
    // offset formatted as `+HH:MM`.
    vm.register_host_fn(
        "wasi:clocks/timezone",
        "to-debug-string",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            if let Some(id) = configured_zone_id() {
                return Value::String(Arc::from(id.as_str()));
            }
            let now = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .as_secs() as i64;
            let text = match local_offset_seconds(now) {
                Some(offset) => {
                    let sign = if offset < 0 { '-' } else { '+' };
                    let magnitude = offset.abs();
                    format!(
                        "{sign}{:02}:{:02}",
                        magnitude / 3600,
                        (magnitude % 3600) / 60
                    )
                }
                None => "no timezone available".to_string(),
            };
            Value::String(Arc::from(text.as_str()))
        }),
    );

    // Thread.Sleep → thread_adapter.rs → wasi:clocks/monotonic-clock.subscribe-duration
    //                                   + wasi:io/poll.[method]pollable.block
    // No `vybe:clocks.sleep` host fn needed — blocking sleep is fully covered by WASI.
}

/// `{ seconds, nanoseconds }` for the current moment — 0.2 `datetime`, 0.3
/// `instant`. Identical shape, so both clocks answer with this.
fn system_clock_now() -> Value {
    let dur = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default();
    let mut rec = Object::new();
    rec.properties
        .insert("seconds".into(), Value::F64(dur.as_secs() as f64));
    rec.properties
        .insert("nanoseconds".into(), Value::F64(dur.subsec_nanos() as f64));
    Value::Object(vybe_runtime::heap::alloc(rec))
}

/// Seconds since the epoch from an `instant`/`datetime` argument.
///
/// Absent or unreadable means "now", which is what a caller asking for the
/// current zone offset means anyway.
fn instant_seconds(when: Option<&Value>) -> i64 {
    if let Some(Value::Object(object)) = when {
        if let Ok(object) = object.lock() {
            if let Some(seconds) = object.properties.get("seconds") {
                return seconds.as_f64() as i64;
            }
        }
    }
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs() as i64
}

/// The configured zone: its IANA identifier and the TZif bytes describing it.
///
/// `TZ` wins when set, matching every POSIX runtime. Otherwise the platform's
/// `/etc/localtime` is followed; when it is a symlink into a zoneinfo tree the
/// suffix IS the IANA identifier, and when it is a plain copy the rules are
/// still readable even though the name is not.
fn configured_zone() -> &'static Option<(Option<String>, Vec<u8>)> {
    static ZONE: OnceLock<Option<(Option<String>, Vec<u8>)>> = OnceLock::new();
    ZONE.get_or_init(|| {
        if let Ok(tz) = std::env::var("TZ") {
            let name = tz.trim_start_matches(':').to_string();
            if !name.is_empty() {
                for root in ["/usr/share/zoneinfo", "/etc/zoneinfo", "/usr/lib/zoneinfo"] {
                    let path = std::path::Path::new(root).join(&name);
                    if let Ok(bytes) = std::fs::read(&path) {
                        return Some((Some(name), bytes));
                    }
                }
                // A `TZ` we cannot resolve to rules still names the zone.
                return Some((Some(name), Vec::new()));
            }
        }

        let localtime = std::path::Path::new("/etc/localtime");
        let bytes = std::fs::read(localtime).ok()?;
        let id = std::fs::read_link(localtime).ok().and_then(|target| {
            let text = target.to_string_lossy().into_owned();
            text.split_once("zoneinfo/")
                .map(|(_, suffix)| suffix.trim_start_matches("posix/").to_string())
        });
        Some((id, bytes))
    })
}

/// The IANA identifier of the configured zone, when the host exposes one.
fn configured_zone_id() -> Option<String> {
    configured_zone().as_ref().and_then(|(id, _)| id.clone())
}

/// Seconds east of UTC at `seconds` since the epoch, per the configured zone's
/// TZif rules. `None` when no zone or no rules are available.
fn local_offset_seconds(seconds: i64) -> Option<i64> {
    tzif_type_at(seconds).map(|(offset, _)| offset)
}

/// Whether the configured zone is in daylight saving time at `seconds`.
fn local_is_dst(seconds: i64) -> Option<bool> {
    tzif_type_at(seconds).map(|(_, is_dst)| is_dst)
}

/// Resolve `(utoff, isdst)` for an instant from the configured zone's TZif.
///
/// RFC 8536: a v1 header and data block, optionally followed for version 2+ by
/// a second header and block whose transition times are 64-bit. The 64-bit
/// block is preferred where present, since the 32-bit one cannot express
/// transitions past 2038.
fn tzif_type_at(seconds: i64) -> Option<(i64, bool)> {
    let (_, bytes) = configured_zone().as_ref()?;
    tzif_type_at_bytes(bytes, seconds)
}

/// The same resolution against explicit TZif bytes.
///
/// Split out so the parser is testable without touching `TZ` — the configured
/// zone is resolved once per process, so a test that set the variable would
/// depend on which test ran first.
pub fn tzif_type_at_bytes(bytes: &[u8], seconds: i64) -> Option<(i64, bool)> {
    {
        if bytes.len() < 44 || &bytes[..4] != b"TZif" {
            return None;
        }
        let version = bytes[4];

        let first = parse_tzif_block(bytes, 0, 4)?;
        let block = if version >= b'2' {
            parse_tzif_block(bytes, first.end, 8).unwrap_or(first)
        } else {
            first
        };

        // Before the first transition, RFC 8536 §3.2 says use the first type that
        // is not daylight saving, falling back to the first type of all.
        let mut chosen = block
            .types
            .iter()
            .find(|(_, is_dst)| !*is_dst)
            .or_else(|| block.types.first())
            .copied()?;
        for (index, transition) in block.transitions.iter().enumerate() {
            if *transition > seconds {
                break;
            }
            if let Some(kind) = block.transition_types.get(index) {
                if let Some(entry) = block.types.get(*kind as usize) {
                    chosen = *entry;
                }
            }
        }
        Some(chosen)
    }
}

struct TzifBlock {
    transitions: Vec<i64>,
    transition_types: Vec<u8>,
    types: Vec<(i64, bool)>,
    end: usize,
}

/// Parse one TZif header + data block starting at `start`, where each
/// transition time occupies `time_size` bytes (4 for v1, 8 for v2+).
fn parse_tzif_block(bytes: &[u8], start: usize, time_size: usize) -> Option<TzifBlock> {
    let header = bytes.get(start..start + 44)?;
    if &header[..4] != b"TZif" {
        return None;
    }
    let count = |index: usize| -> usize {
        let at = 20 + index * 4;
        u32::from_be_bytes([header[at], header[at + 1], header[at + 2], header[at + 3]]) as usize
    };
    let (isutcnt, isstdcnt, leapcnt, timecnt, typecnt, charcnt) =
        (count(0), count(1), count(2), count(3), count(4), count(5));

    let mut at = start + 44;
    let mut transitions = Vec::with_capacity(timecnt);
    for _ in 0..timecnt {
        let raw = bytes.get(at..at + time_size)?;
        transitions.push(match time_size {
            8 => i64::from_be_bytes(raw.try_into().ok()?),
            _ => i32::from_be_bytes(raw.try_into().ok()?) as i64,
        });
        at += time_size;
    }

    let transition_types = bytes.get(at..at + timecnt)?.to_vec();
    at += timecnt;

    let mut types = Vec::with_capacity(typecnt);
    for _ in 0..typecnt {
        let raw = bytes.get(at..at + 6)?;
        let utoff = i32::from_be_bytes([raw[0], raw[1], raw[2], raw[3]]) as i64;
        types.push((utoff, raw[4] != 0));
        at += 6;
    }

    // Designations, leap seconds and the standard/UT indicators are not needed
    // to answer an offset, but their sizes place the next block.
    at += charcnt + leapcnt * (time_size + 4) + isstdcnt + isutcnt;
    Some(TzifBlock {
        transitions,
        transition_types,
        types,
        end: at,
    })
}
