use super::helpers::run_prints;

#[test]
fn test_datetime_create_from_format_reset_fields() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat('!Y-m-d', '2024-06-15', new DateTimeZone('UTC'));
echo $dt->format('Y-m-d H:i:s'), "\n";
"#
        ),
        vec!["2024-06-15 00:00:00"]
    );
}

#[test]
fn test_datetime_create_from_format_invalid_errors() {
    assert_eq!(
        run_prints(
            r#"<?php
$res = DateTime::createFromFormat('Y-m-d', 'invalid-date');
$errors = DateTime::getLastErrors();
echo ($res === false && is_array($errors) && ($errors['error_count'] > 0 || $errors['warning_count'] > 0 || count($errors['errors']) > 0)) ? 'errors_logged' : 'err', "\n";
"#
        ),
        vec!["errors_logged"]
    );
}

#[test]
fn test_datetime_create_from_format_unix_timestamp() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat('U', '1718449800', new DateTimeZone('UTC'));
echo $dt->format('Y-m-d H:i:s'), "\n";
"#
        ),
        vec!["2024-06-15 11:10:00"]
    );
}

#[test]
fn test_datetime_create_from_format_microseconds() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat('Y-m-d H:i:s.u', '2024-12-31 23:59:59.123456', new DateTimeZone('UTC'));
echo $dt !== false ? $dt->format('u') : 'bad';
"#
        ),
        vec!["123456"]
    );
}

#[test]
fn test_datetime_create_from_format_trailing_space() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat('Y-m-d', '2024-01-01 ');
$errors = DateTime::getLastErrors();
echo $dt !== false ? 'ok' : 'bad';
echo '|';
echo (($errors['warning_count'] ?? 0) > 0) ? 'warn' : 'nowarn';
"#
        ),
        vec!["ok|warn"]
    );
}

#[test]
fn test_datetime_create_from_format_timezone_alias() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat('Y-m-d H:i:s P', '2024-03-01 12:00:00 -0500');
echo $dt !== false ? $dt->format('P') : 'bad';
"#
        ),
        vec!["-05:00"]
    );
}
