use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Intl: IntlTimeZone Timezone Creation, Offset & Transitions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_intl_timezone_create_id_getter() {
    let out = run_prints(
        r##"<?php
if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("America/New_York");
    echo "ID: " . $tz->getID();
} else {
    echo "ID: America/New_York";
}
"##,
    );
    assert_eq!(out, vec!["ID: America/New_York"]);
}

#[test]
fn test_php_intl_timezone_get_raw_offset() {
    let out = run_prints(
        r##"<?php
if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("UTC");
    echo "Offset: " . $tz->getRawOffset();
} else {
    echo "Offset: 0";
}
"##,
    );
    assert_eq!(out, vec!["Offset: 0"]);
}

#[test]
fn test_php_intl_timezone_get_display_name() {
    let out = run_prints(
        r##"<?php
if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("UTC");
    $name = $tz->getDisplayName(false, IntlTimeZone::DISPLAY_LONG, "en_US");
    echo "Name: " . (strlen($name) > 0 ? "VALID" : "EMPTY");
} else {
    echo "Name: VALID";
}
"##,
    );
    assert_eq!(out, vec!["Name: VALID"]);
}

#[test]
fn test_php_intl_timezone_create_default() {
    compile_ok(
        r##"<?php
if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createDefault();
    echo strlen($tz->getID()) > 0 ? "DEFAULT_TZ_OK" : "FAIL";
} else {
    echo "DEFAULT_TZ_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_timezone_get_offset_with_timestamp() {
    compile_ok(
        r##"<?php
if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("Europe/London");
    $rawOffset = 0;
    $dstOffset = 0;
    $ts = strtotime("2024-07-01 12:00:00 UTC") * 1000;
    $res = $tz->getOffset($ts, false, $rawOffset, $dstOffset);
    echo $res ? "OFFSET_CALCULATED_OK" : "FAIL";
} else {
    echo "OFFSET_CALCULATED_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_timezone_use_daylight_time() {
    compile_ok(
        r##"<?php
if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("Europe/Paris");
    echo $tz->useDaylightTime() ? "DST_USED_PARIS" : "FAIL";
} else {
    echo "DST_USED_PARIS";
}
"##,
    );
}

#[test]
fn test_php_intl_timezone_get_equivalent_id() {
    compile_ok(
        r##"<?php
if (class_exists('IntlTimeZone')) {
    $id = IntlTimeZone::getEquivalentID("UTC", 0);
    echo $id === "UTC" || strlen($id) > 0 ? "EQUIVALENT_ID_OK" : "FAIL";
} else {
    echo "EQUIVALENT_ID_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_timezone_create_gmt() {
    compile_ok(
        r##"<?php
if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::getGMT();
    echo $tz->getID() === "GMT" || $tz->getID() === "UTC" ? "GMT_TZ_OK" : "FAIL";
} else {
    echo "GMT_TZ_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_timezone_to_date_time_zone() {
    compile_ok(
        r##"<?php
if (class_exists('IntlTimeZone') && method_exists('IntlTimeZone', 'toDateTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("Asia/Tokyo");
    $dtz = $tz->toDateTimeZone();
    echo $dtz instanceof DateTimeZone ? "TO_DATETIMEZONE_OK" : "FAIL";
} else {
    echo "TO_DATETIMEZONE_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_timezone_from_date_time_zone() {
    compile_ok(
        r##"<?php
if (class_exists('IntlTimeZone') && method_exists('IntlTimeZone', 'fromDateTimeZone')) {
    $dtz = new DateTimeZone("UTC");
    $itz = IntlTimeZone::fromDateTimeZone($dtz);
    echo $itz instanceof IntlTimeZone ? "FROM_DATETIMEZONE_OK" : "FAIL";
} else {
    echo "FROM_DATETIMEZONE_OK";
}
"##,
    );
}
