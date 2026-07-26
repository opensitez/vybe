use super::helpers::run_prints;

#[test]
fn test_datetime_immutable_create_from_format_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTimeImmutable::createFromFormat('Y/m/d H:i', '2024/08/20 14:45', new DateTimeZone('UTC'));
echo $dt->format('Y-m-d H:i:s') . ':' . get_class($dt), "\n";
"#
        ),
        vec!["2024-08-20 14:45:00:DateTimeImmutable"]
    );
}

#[test]
fn test_datetime_immutable_create_from_format_reset_time() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTimeImmutable::createFromFormat('!d-m-Y', '25-12-2024', new DateTimeZone('UTC'));
echo $dt->format('Y-m-d H:i:s'), "\n";
"#
        ),
        vec!["2024-12-25 00:00:00"]
    );
}

#[test]
fn test_datetime_immutable_create_from_format_error_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTimeImmutable::createFromFormat('Y-m-d', 'not-a-date');
echo $dt === false ? 'false' : 'ok';
$e = DateTimeImmutable::getLastErrors();
echo '|';
echo (($e['error_count'] ?? 0) > 0) ? 'errs' : 'clean';
"#
        ),
        vec!["false|errs"]
    );
}

#[test]
fn test_datetime_immutable_create_from_format_unix_timestamp() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTimeImmutable::createFromFormat('U', '1704067200', new DateTimeZone('UTC'));
echo $dt->format('Y-m-d H:i:s'), "\n";
"#
        ),
        vec!["2024-01-01 00:00:00"]
    );
}

#[test]
fn test_datetime_immutable_create_from_format_timezone_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTimeImmutable::createFromFormat('Y-m-d H:i P', '2024-01-01 12:00 -0500', new DateTimeZone('UTC'));
echo $dt !== false ? $dt->format('P') : 'bad';
"#
        ),
        vec!["-05:00"]
    );
}

#[test]
fn test_datetime_immutable_create_from_format_microseconds() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTimeImmutable::createFromFormat('Y-m-d H:i:s.u', '2024-12-31 23:59:59.999999', new DateTimeZone('UTC'));
echo $dt !== false ? $dt->format('u') : 'bad';
"#
        ),
        vec!["999999"]
    );
}
