use super::helpers::run_prints;

fn assert_output(expr: &str, expected: &str) {
    assert_eq!(run_prints(&format!("<?php echo {}; ", expr)), vec![expected.to_string()]);
}

fn assert_int(expr: &str, expected: i64) {
    assert_output(expr, &expected.to_string());
}

#[test]
fn php_datetime_and_timezones() {
    let zones = [
        "UTC",
        "Europe/London",
        "Europe/Berlin",
        "America/New_York",
        "America/Los_Angeles",
        "Asia/Tokyo",
        "Asia/Kolkata",
        "Australia/Sydney",
        "America/Sao_Paulo",
        "Africa/Casablanca",
    ];

    let base = 1_700_000_000_i64;
    for idx in 0..10_i64 {
        let ts = base + (idx * 86_400);
        assert_int(&format!("(new DateTimeImmutable('@{ts}'))->format('U')"), ts);
    }

    for idx in 0..10_i64 {
        let zone = zones[idx as usize];
        assert_output(
            &format!(
                "(new DateTimeImmutable('2024-01-01 00:00:00', new DateTimeZone('{}')))->format('e')",
                zone
            ),
            zone,
        );
    }
}

#[test]
fn php_datetime_timezone_offsets_and_transitions() {
    assert_output(
        "(new DateTimeZone('UTC'))->getOffset(new DateTimeImmutable('2024-01-01 00:00:00', new DateTimeZone('UTC')))",
        "0",
    );

    assert_output(
        "(new DateTimeImmutable('2024-01-01 12:00:00', new DateTimeZone('America/New_York')))->setTimezone(new DateTimeZone('Europe/London'))->format('e')",
        "Europe/London",
    );
}

#[test]
fn php_datetime_timezone_abbreviations_and_transition_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Europe/Berlin');
$abbrevs = DateTimeZone::listAbbreviations();
echo array_key_exists('ce', $abbrevs) ? 'ce' : 'no';
$transitions = $tz->getTransitions(strtotime('2024-01-01'), strtotime('2024-07-01'));
echo count($transitions) > 0 ? 'has' : 'none';
"#
        ),
        vec!["ceno"]
    );
}

#[test]
fn php_datetime_timezone_locations_and_country() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Africa/Casablanca');
$loc = $tz->getLocation();
echo $loc['country_code'];
echo $loc['timezone_id'];
"#,
        ),
        vec!["MAAfrica/Casablanca"]
    );
}

#[test]
fn php_datetime_timezone_list_filtering_by_group() {
    assert_eq!(
        run_prints(
            r#"<?php
$tzs = DateTimeZone::listIdentifiers(DateTimeZone::AMERICA);
echo is_array($tzs) ? 'array' : 'no';
echo '|';
echo in_array('America/New_York', $tzs, true) ? 'ny' : 'no';
"#,
        ),
        vec!["array|ny"]
    );
}

#[test]
fn php_datetime_timezone_immutability_after_set_timezone() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = new DateTime('2024-07-01 12:00:00', new DateTimeZone('UTC'));
$next = new DateTime('2024-07-01 12:00:00', new DateTimeZone('America/New_York'));
$next->setTimezone(new DateTimeZone('UTC'));
echo $base->format('H');
echo $next->format('H');
"#,
        ),
        vec!["12|12"]
    );
}

#[test]
fn php_datetime_timezone_offset_switch_across_dst_boundary() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('America/New_York');
$winter = new DateTime('2024-01-15 12:00:00', $tz);
$summer = new DateTime('2024-07-15 12:00:00', $tz);
echo $tz->getOffset($winter);
echo '|';
echo $tz->getOffset($summer);
"#,
        ),
        vec!["-18000|-14400"]
    );
}

#[test]
fn php_datetime_default_timezone_switch_roundtrip() {
    assert_eq!(
        run_prints(
            r#"<?php
$before = date_default_timezone_get();
date_default_timezone_set('UTC');
$utc = date_default_timezone_get();
date_default_timezone_set('America/Los_Angeles');
$la = date_default_timezone_get();
date_default_timezone_set($before);
echo ($utc === 'UTC') ? 'utc' : 'not';
echo '|';
echo ($la === 'America/Los_Angeles') ? 'la' : 'no';
"#,
        ),
        vec!["utc|la"]
    );
}

#[test]
fn php_datetime_timezone_settimestamp_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-01-01 00:00:00", new DateTimeZone('UTC'));
$dt->setTimestamp(1710000000);
echo $dt->format('Y-m-d');
"#,
        ),
        vec!["2024-03-09"]
    );
}

#[test]
fn php_datetime_timezone_construct_invalid_zone_errors() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    new DateTimeZone('Not/AZone');
    echo 'ok';
} catch (Throwable $e) {
    echo 'err';
}
"#,
        ),
        vec!["err"]
    );
}

#[test]
fn php_datetime_timezone_get_offset_for_fixed_timestamp() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Europe/Paris');
$dt = new DateTimeImmutable('@1710000000', $tz);
echo $tz->getOffset($dt);
"#,
        ),
        vec!["3600"]
    );
}

#[test]
fn php_datetime_timezone_list_identifiers_asia_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::ASIA);
echo is_array($zones) ? 'array' : 'bad';
echo '|';
echo in_array('Asia/Tokyo', $zones, true) ? 'has_tokyo' : 'no_tokyo';
"#,
        ),
        vec!["array|has_tokyo"]
    );
}

#[test]
fn php_datetime_timezone_list_identifiers_country() {
    assert_eq!(
        run_prints(
            r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::PER_COUNTRY, 'IN');
echo is_array($zones) ? 'array' : 'bad';
echo '|';
echo in_array('Asia/Kolkata', $zones, true) ? 'kolkata' : 'no';
"#
        ),
        vec!["array|kolkata"]
    );
}
