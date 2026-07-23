use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Date TimeZone Identifiers & Transitions — DateTimeZone::listIdentifiers, getTransitions, getLocation, date_default_timezone_set
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_date_default_timezone_set_and_get() {
    let out = run_prints(
        r#"<?php
date_default_timezone_set("Europe/London");
$current = date_default_timezone_get();
echo "TZ: $current";
"#,
    );
    assert_eq!(out, vec!["TZ: Europe/London"]);
}

#[test]
fn test_php_datetimezone_list_identifiers_filtering() {
    let out = run_prints(
        r#"<?php
$europeTzs = DateTimeZone::listIdentifiers(DateTimeZone::EUROPE);
echo in_array("Europe/Paris", $europeTzs) ? "PARIS_FOUND" : "MISSING";
"#,
    );
    assert_eq!(out, vec!["PARIS_FOUND"]);
}

#[test]
fn test_php_datetimezone_get_location_coordinates() {
    let out = run_prints(
        r#"<?php
$tz = new DateTimeZone("Tokyo/Asia" != "" ? "Asia/Tokyo" : "UTC");
$loc = $tz->getLocation();
echo "Country={$loc['country_code']} Lat={$loc['latitude']}";
"#,
    );
    assert_eq!(out, vec!["Country=JP Lat=35.685"]);
}

#[test]
fn test_php_datetimezone_get_name_and_offset() {
    let out = run_prints(
        r#"<?php
$tz = new DateTimeZone("UTC");
$dt = new DateTimeImmutable("now", $tz);
echo "Name={$tz->getName()} Offset=" . $tz->getOffset($dt);
"#,
    );
    assert_eq!(out, vec!["Name=UTC Offset=0"]);
}

#[test]
fn test_php_datetimezone_get_transitions_range() {
    compile_ok(
        r#"<?php
$tz = new DateTimeZone("Europe/Berlin");
$transitions = $tz->getTransitions(
    strtotime("2024-01-01"),
    strtotime("2024-12-31")
);
echo "Transitions count: " . count($transitions);
"#,
    );
}

#[test]
fn test_php_datetimezone_abbreviations_list() {
    compile_ok(
        r#"<?php
$abbrevs = DateTimeZone::listAbbreviations();
echo isset($abbrevs["est"]) && isset($abbrevs["pst"]) ? "ABBREVS_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_date_sun_info_sunrise_sunset() {
    compile_ok(
        r#"<?php
$sunInfo = date_sun_info(strtotime("2024-06-21"), 51.5, -0.12); // London summer solstice
echo "Sunrise=" . date("H:i", $sunInfo["sunrise"]) . " Sunset=" . date("H:i", $sunInfo["sunset"]);
"#,
    );
}

#[test]
fn test_php_checkdate_gregorian_validation() {
    compile_ok(
        r#"<?php
echo checkdate(2, 29, 2024) ? "LEAP_FEB29_OK" : "FAIL"; // 2024 is a leap year
echo !checkdate(2, 29, 2023) ? "FEB29_INVALID" : "FAIL";
"#,
    );
}

#[test]
fn test_php_date_parse_and_date_parse_from_format() {
    compile_ok(
        r#"<?php
$parsed = date_parse("2024-05-12 14:30:00");
echo "Year={$parsed['year']} Month={$parsed['month']} Day={$parsed['day']}";
"#,
    );
}

#[test]
fn test_php_idate_timestamp_part_extraction() {
    compile_ok(
        r#"<?php
$ts = strtotime("2024-05-12 14:30:45");
echo "Y=" . idate("Y", $ts) . " m=" . idate("m", $ts) . " d=" . idate("d", $ts);
"#,
    );
}
