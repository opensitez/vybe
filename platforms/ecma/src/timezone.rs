//! IANA time zone support — the data layer ECMA-402 §"Use of the IANA Time
//! Zone Database" requires.
//!
//! The spec is explicit that it does NOT carry this data itself: the IANA Time
//! Zone Database is a *normative reference*, implementations "must be time zone
//! aware: they must use the IANA Time Zone Database to supply available named
//! time zone identifiers and data used in ECMAScript calculations and
//! formatting", and it notes tzdb is "typically updated between five and ten
//! times per year". So the database is linked (`chrono-tz`, which compiles
//! tzdata at build time), never hand-tabulated.
//!
//! These operations back ECMA-402's `AvailableNamedTimeZoneIdentifiers` and
//! `IsValidTimeZoneName`, and they are deliberately language-neutral: PHP's
//! `DateTimeZone::getOffset`/`getTransitions` and Java's `ZoneId` want exactly
//! the same four answers, and previously each carried its own stub — a
//! name-passthrough in `instant_adapter.rs`, a hardcoded `"US"` in PHP's
//! `datetime_adapter.rs`, a 4-entry short-id map, and a `RangeError` here.

use chrono::{Offset, TimeZone as _, Utc};
use chrono_tz::{TZ_VARIANTS, Tz};
use std::str::FromStr;
use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

/// Resolve a time zone identifier to its tzdb zone.
///
/// Per spec, `"UTC"` is a primary identifier "for historical reasons", and
/// `"Etc/UTC"`, `"Etc/GMT"` and `"GMT"` resolve to it. Lookup is
/// case-insensitive on the way in because callers spell zones as the user
/// typed them, but every identifier this module HANDS BACK uses tzdb casing,
/// which the spec requires.
pub fn resolve(name: &str) -> Option<Tz> {
    let trimmed = name.trim();
    if trimmed.is_empty() {
        return None;
    }
    if trimmed.eq_ignore_ascii_case("UTC")
        || trimmed.eq_ignore_ascii_case("GMT")
        || trimmed.eq_ignore_ascii_case("Etc/UTC")
        || trimmed.eq_ignore_ascii_case("Etc/GMT")
        || trimmed == "Z"
    {
        return Some(Tz::UTC);
    }
    if let Ok(tz) = Tz::from_str(trimmed) {
        return Some(tz);
    }
    // Case-insensitive fallback so `america/new_york` still resolves; the name
    // returned to callers is the tzdb spelling, not the caller's.
    TZ_VARIANTS
        .iter()
        .find(|tz| tz.name().eq_ignore_ascii_case(trimmed))
        .copied()
}

/// ECMA-402 `IsValidTimeZoneName`.
pub fn is_valid(name: &str) -> bool {
    resolve(name).is_some()
}

/// The canonical (tzdb-cased) identifier, with the spec's `"UTC"` exception.
pub fn canonicalize(name: &str) -> Option<String> {
    let tz = resolve(name)?;
    Some(if tz == Tz::UTC {
        "UTC".to_string()
    } else {
        tz.name().to_string()
    })
}

/// Offset from UTC in SECONDS EAST at the given instant — the sign convention
/// of tzdb and of PHP's `DateTimeZone::getOffset`. Note this is the opposite
/// sign from JavaScript's `Date.prototype.getTimezoneOffset`, which reports
/// minutes WEST; converting is the caller's job so the difference stays
/// visible at the call site.
pub fn offset_seconds(name: &str, ms: f64) -> Option<i32> {
    let tz = resolve(name)?;
    let dt = Utc.timestamp_millis_opt(ms as i64).single()?;
    Some(tz.offset_from_utc_datetime(&dt.naive_utc()).fix().local_minus_utc())
}

/// Whether daylight saving is in effect for `name` at `ms`.
///
/// Derived by comparing the zone's offset at the instant against its offset in
/// January and July of the same year: DST is the LARGER of the two seasonal
/// offsets, which works for both hemispheres without a per-region table.
pub fn is_dst(name: &str, ms: f64) -> bool {
    let Some(current) = offset_seconds(name, ms) else {
        return false;
    };
    let Some(dt) = Utc.timestamp_millis_opt(ms as i64).single() else {
        return false;
    };
    let year_start = Utc
        .with_ymd_and_hms(dt.format("%Y").to_string().parse().unwrap_or(1970), 1, 15, 0, 0, 0)
        .single();
    let mid_year = Utc
        .with_ymd_and_hms(dt.format("%Y").to_string().parse().unwrap_or(1970), 7, 15, 0, 0, 0)
        .single();
    let (Some(jan), Some(jul)) = (year_start, mid_year) else {
        return false;
    };
    let jan_off = offset_seconds(name, jan.timestamp_millis() as f64).unwrap_or(current);
    let jul_off = offset_seconds(name, jul.timestamp_millis() as f64).unwrap_or(current);
    if jan_off == jul_off {
        return false; // Zone observes no DST at all.
    }
    current == jan_off.max(jul_off)
}

/// Zone abbreviation in effect at an instant (`EST`, `BST`, `JST`, …).
pub fn abbreviation(name: &str, ms: f64) -> Option<String> {
    let tz = resolve(name)?;
    let dt = Utc.timestamp_millis_opt(ms as i64).single()?;
    Some(tz.from_utc_datetime(&dt.naive_utc()).format("%Z").to_string())
}

/// IANA `zone.tab` — the ISO 3166-1 alpha-2 country code → zone identifier
/// table. ECMA-402 names this file directly ("Any Link name that is present in
/// the 'TZ' column of file zone.tab must be a primary time zone identifier"),
/// and it is the ONLY source of the region mapping `Intl.Locale.getTimeZones`
/// needs: `chrono-tz` vendors the tzdb rule files but not this table, and
/// ICU's region data covers only Windows-zone disambiguation.
///
/// Public domain, per the tzdb LICENSE.
const ZONE_TAB: &str = include_str!("../data/zone.tab");

/// Zones for an ISO 3166-1 alpha-2 region, tzdb casing, sorted.
///
/// Entries are validated against `TZ_VARIANTS` because this table and
/// `chrono-tz`'s rules are separate tzdb releases (2026b vs 2025b at time of
/// writing). Without the check, a zone added in the newer release would be
/// returned here and then fail to resolve in `offset_seconds` — the same
/// module giving inconsistent answers.
pub fn identifiers_for_region(region: &str) -> Vec<String> {
    if region.len() != 2 || !region.chars().all(|c| c.is_ascii_alphabetic()) {
        return Vec::new();
    }
    let mut out: Vec<String> = ZONE_TAB
        .lines()
        .filter(|line| !line.starts_with('#') && !line.trim().is_empty())
        .filter_map(|line| {
            let mut cols = line.split('\t');
            let codes = cols.next()?;
            let _coordinates = cols.next()?;
            let zone = cols.next()?.trim();
            // zone.tab is one country per row; zone1970.tab uses a
            // comma-separated list, so accept both shapes.
            codes
                .split(',')
                .any(|code| code.eq_ignore_ascii_case(region))
                .then(|| zone.to_string())
        })
        .filter(|zone| resolve(zone).is_some())
        .collect();
    out.sort();
    out.dedup();
    out
}

/// VM global holding the host environment's current time zone — the ONE clock
/// every layer reads.
///
/// ECMA-262 `SystemTimeZoneIdentifier()` is specified as "a String representing
/// **the host environment's** current time zone", and JavaScript deliberately
/// offers no way to set it (you change `TZ`, not a JS property). PHP
/// (`date_default_timezone_set`), Java (`TimeZone.setDefault`) and .NET
/// (`TimeZoneInfo`) all DO expose a setter — so the value is host-environment
/// state that some languages may write and every language reads.
///
/// It therefore lives in one VM global rather than in any language's adapter:
/// set it from PHP and Java, Python, .NET and `Intl` all observe it.
pub const DEFAULT_TZ_GLOBAL: &str = "__vybe_system_timezone";

/// The one clock. Process-level because the value it models IS process state —
/// the same thing `TZ` is for a Unix process — so every VM, every language
/// adapter and every host module observes one setting. A per-VM global would
/// let PHP and Java disagree, which is the bug this exists to prevent.
static SYSTEM_TZ: std::sync::OnceLock<std::sync::RwLock<String>> = std::sync::OnceLock::new();

fn system_tz_cell() -> &'static std::sync::RwLock<String> {
    SYSTEM_TZ.get_or_init(|| std::sync::RwLock::new("UTC".to_string()))
}

/// Set the host environment's zone. Returns false for an identifier tzdb does
/// not know. Stores the CANONICAL spelling, so reads always yield a primary
/// identifier as ECMA-262 requires.
pub fn set_system_identifier(name: &str) -> bool {
    match canonicalize(name) {
        Some(canonical) => {
            if let Ok(mut guard) = system_tz_cell().write() {
                *guard = canonical;
                return true;
            }
            false
        }
        None => false,
    }
}

/// ECMA-262 `SystemTimeZoneIdentifier()` — the host environment's zone as a
/// PRIMARY identifier, defaulting to `"UTC"`.
///
/// The spec permits returning `"UTC"` unconditionally only "if the
/// implementation only supports the UTC time zone". Once tzdb is linked that
/// exemption no longer applies, so this must reflect the real setting and
/// `GetNamedTimeZoneOffsetNanoseconds` must agree with it.
pub fn system_identifier() -> String {
    system_tz_cell()
        .read()
        .map(|guard| guard.clone())
        .unwrap_or_else(|_| "UTC".to_string())
}

/// Offset of the host environment's zone at an instant, in seconds EAST.
pub fn system_offset_seconds(ms: f64) -> i32 {
    offset_seconds(&system_identifier(), ms).unwrap_or(0)
}

fn s(value: &str) -> Value {
    Value::String(Arc::from(value))
}

fn arg_str(args: &[Value], idx: usize) -> Option<String> {
    match args.get(idx) {
        Some(Value::String(v)) => Some(v.to_string()),
        _ => None,
    }
}

fn arg_ms(args: &[Value], idx: usize) -> f64 {
    match args.get(idx) {
        Some(Value::F64(v)) => *v,
        Some(Value::I32(v)) => *v as f64,
        _ => 0.0,
    }
}

fn make_array(items: Vec<Value>) -> Value {
    let mut obj = Object::new_array(items);
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Array")));
    obj.properties
        .insert("__proto__".into(), crate::array::shared_array_prototype());
    Value::Object(vybe_runtime::heap::alloc(obj))
}

pub fn register(vm: &mut VM) {
    // AvailableNamedTimeZoneIdentifiers — every Zone/Link name tzdb knows,
    // in tzdb casing, sorted. Replaces the hardcoded four-entry lists that
    // each language carried.
    vm.register_host_fn(
        "ecma:intl/timezone",
        "identifiers",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut names: Vec<&str> = TZ_VARIANTS.iter().map(|tz| tz.name()).collect();
            names.sort_unstable();
            make_array(names.into_iter().map(s).collect())
        }),
    );

    // IsValidTimeZoneName
    vm.register_host_fn(
        "ecma:intl/timezone",
        "isValid",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(arg_str(args, 0).map(|n| is_valid(&n)).unwrap_or(false))
        }),
    );

    // CanonicalizeTimeZoneName — null when the identifier is not in tzdb, so
    // callers can distinguish "unknown zone" from "resolved to itself".
    vm.register_host_fn(
        "ecma:intl/timezone",
        "canonicalize",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match arg_str(args, 0).and_then(|n| canonicalize(&n)) {
                Some(name) => s(&name),
                None => Value::Null,
            }
        }),
    );

    // offset(name, msSinceEpoch) → seconds EAST of UTC.
    vm.register_host_fn(
        "ecma:intl/timezone",
        "offset",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match arg_str(args, 0).and_then(|n| offset_seconds(&n, arg_ms(args, 1))) {
                Some(secs) => Value::I32(secs),
                None => Value::Null,
            }
        }),
    );

    // ECMA-262 SystemTimeZoneIdentifier() — read the one clock.
    vm.register_host_fn(
        "ecma:intl/timezone",
        "systemIdentifier",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| s(&system_identifier())),
    );

    // Set the host environment's zone. NOT a JavaScript operation — JS changes
    // this via `TZ`, not an API — but PHP/Java/.NET expose setters, and they
    // must all write the value `SystemTimeZoneIdentifier` reads. Rejects an
    // identifier tzdb does not know, and stores the CANONICAL spelling so the
    // read side always returns a primary identifier.
    vm.register_host_fn(
        "ecma:intl/timezone",
        "setSystemIdentifier",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| match arg_str(args, 0) {
            Some(name) => Value::Bool(set_system_identifier(&name)),
            None => Value::Bool(false),
        }),
    );

    // isDst(name, msSinceEpoch) — whether daylight saving is in effect.
    vm.register_host_fn(
        "ecma:intl/timezone",
        "isDst",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(name) = arg_str(args, 0) else {
                return Value::Bool(false);
            };
            Value::Bool(is_dst(&name, arg_ms(args, 1)))
        }),
    );

    // abbreviation(name, msSinceEpoch) → "EST" / "BST" / …
    vm.register_host_fn(
        "ecma:intl/timezone",
        "abbreviation",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match arg_str(args, 0).and_then(|n| abbreviation(&n, arg_ms(args, 1))) {
                Some(abbr) => s(&abbr),
                None => Value::Null,
            }
        }),
    );
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 2024-01-01T00:00:00Z (winter) and 2024-07-01T00:00:00Z (summer) —
    /// the pair that distinguishes a real tzdb from a fixed-offset table.
    const WINTER_MS: f64 = 1_704_067_200_000.0;
    const SUMMER_MS: f64 = 1_719_792_000_000.0;

    #[test]
    fn dst_offsets_differ_across_the_year() {
        assert_eq!(offset_seconds("America/New_York", WINTER_MS), Some(-18000));
        assert_eq!(offset_seconds("America/New_York", SUMMER_MS), Some(-14400));
        assert_eq!(offset_seconds("Europe/London", WINTER_MS), Some(0));
        assert_eq!(offset_seconds("Europe/London", SUMMER_MS), Some(3600));
    }

    #[test]
    fn half_hour_and_southern_hemisphere_zones() {
        // Kolkata is +05:30 year-round; Sydney's DST runs the opposite way.
        assert_eq!(offset_seconds("Asia/Kolkata", WINTER_MS), Some(19800));
        assert_eq!(offset_seconds("Asia/Kolkata", SUMMER_MS), Some(19800));
        assert_eq!(offset_seconds("Australia/Sydney", WINTER_MS), Some(39600));
        assert_eq!(offset_seconds("Australia/Sydney", SUMMER_MS), Some(36000));
    }

    #[test]
    fn utc_is_primary_and_links_resolve_to_it() {
        for alias in ["UTC", "Etc/UTC", "Etc/GMT", "GMT", "Z"] {
            assert_eq!(canonicalize(alias).as_deref(), Some("UTC"), "{alias}");
        }
    }

    #[test]
    fn identifiers_use_tzdb_casing_and_invalid_zones_are_rejected() {
        assert_eq!(
            canonicalize("america/new_york").as_deref(),
            Some("America/New_York")
        );
        assert!(!is_valid("Not/AZone"));
        assert_eq!(canonicalize("Not/AZone"), None);
    }

    #[test]
    fn abbreviations_track_daylight_saving() {
        assert_eq!(abbreviation("America/New_York", WINTER_MS).as_deref(), Some("EST"));
        assert_eq!(abbreviation("America/New_York", SUMMER_MS).as_deref(), Some("EDT"));
    }
}
