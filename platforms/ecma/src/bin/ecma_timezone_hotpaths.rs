//! Native timezone lookup equivalence and timing; not VM/application timing.
use chrono_tz::{TZ_VARIANTS, Tz};
use std::{hint::black_box, str::FromStr, time::Instant};
use vybe_platform_ecma::timezone;

fn original_resolve(name: &str) -> Option<Tz> {
    let name = name.trim();
    if name.is_empty() {
        return None;
    }
    if name.eq_ignore_ascii_case("UTC")
        || name.eq_ignore_ascii_case("GMT")
        || name.eq_ignore_ascii_case("Etc/UTC")
        || name.eq_ignore_ascii_case("Etc/GMT")
        || name == "Z"
    {
        return Some(Tz::UTC);
    }
    Tz::from_str(name).ok().or_else(|| {
        TZ_VARIANTS.iter().find(|zone| zone.name().eq_ignore_ascii_case(name)).copied()
    })
}

fn main() {
    for zone in TZ_VARIANTS.iter() {
        for input in [zone.name().to_owned(), zone.name().to_ascii_lowercase(),
            zone.name().to_ascii_uppercase(), format!("  {}\t", zone.name())] {
            assert_eq!(timezone::resolve(&input), original_resolve(&input), "{input:?}");
        }
    }
    for input in ["", " ", "UTC", "utc", "GMT", "etc/gmt", "Z", "z",
        "america/not_a_zone", "not_a_zone", "Europe/Paris/extra", "\u{0130}",
        "\u{00a0}America/New_York\u{00a0}"] {
        assert_eq!(timezone::resolve(input), original_resolve(input), "{input:?}");
    }
    let oversized = "x".repeat(1_000_000);
    assert_eq!(timezone::resolve(&oversized), original_resolve(&oversized));
    // Resolution caching must not cache an offset across seasonal transitions.
    assert_eq!(timezone::offset_seconds("america/new_york", 1705320000000.0), Some(-18000));
    assert_eq!(timezone::offset_seconds("america/new_york", 1721044800000.0), Some(-14400));
    println!("tzdb exact/case-folded/alias/invalid/oversized equivalence and seasonal offsets passed");
    for input in ["America/New_York", "america/new_york", "not_a_zone"] {
        timezone::resolve(input);
        let start = Instant::now();
        for _ in 0..10000 {
            black_box(original_resolve(black_box(input)));
        }
        let original = start.elapsed();
        let start = Instant::now();
        for _ in 0..10000 {
            black_box(timezone::resolve(black_box(input)));
        }
        println!("{input}: 10000 native lookups original={original:?}, indexed={:?}", start.elapsed());
    }
}
