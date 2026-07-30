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

/// Zone abbreviation in effect at an instant (`EST`, `BST`, `JST`, …).
pub fn abbreviation(name: &str, ms: f64) -> Option<String> {
    let tz = resolve(name)?;
    let dt = Utc.timestamp_millis_opt(ms as i64).single()?;
    Some(tz.from_utc_datetime(&dt.naive_utc()).format("%Z").to_string())
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
