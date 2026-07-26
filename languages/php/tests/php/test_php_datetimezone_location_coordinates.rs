use super::helpers::run_prints;

#[test]
fn test_datetimezone_get_location_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Europe/London');
$loc = $tz->getLocation();
echo is_array($loc) && isset($loc['country_code']) && isset($loc['latitude']) ? 'location_ok' : 'err', "\n";
"#
        ),
        vec!["location_ok"]
    );
}

#[test]
fn test_datetimezone_get_location_country_code() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('America/New_York');
$loc = $tz->getLocation();
echo $loc['country_code'], "\n";
"#
        ),
        vec!["US"]
    );
}

#[test]
fn test_datetimezone_get_location_has_longitude() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Asia/Tokyo');
$loc = $tz->getLocation();
echo is_numeric($loc['longitude']) ? 'lon' : 'nolon';
"#
        ),
        vec!["lon"]
    );
}

#[test]
fn test_datetimezone_get_location_has_comments() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Europe/Paris');
$loc = $tz->getLocation();
echo isset($loc['comments']) ? 'comments_key' : 'no_comments_key';
"#
        ),
        vec!["comments_key"]
    );
}

#[test]
fn test_datetimezone_get_location_country_code_is_string() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Africa/Cairo');
$loc = $tz->getLocation();
echo is_string($loc['country_code']) ? 'cc_string' : 'cc_not_string';
"#
        ),
        vec!["cc_string"]
    );
}
