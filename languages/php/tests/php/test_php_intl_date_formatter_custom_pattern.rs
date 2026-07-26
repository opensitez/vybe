use super::helpers::run_prints;

#[test]
fn test_intl_date_formatter_custom_pattern_format() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('IntlDateFormatter')) {
    $fmt = new IntlDateFormatter('en_US', IntlDateFormatter::FULL, IntlDateFormatter::FULL, 'UTC', IntlDateFormatter::GREGORIAN, 'yyyy-MM-dd HH:mm:ss');
    $dt = new DateTime('2024-05-15 10:20:30', new DateTimeZone('UTC'));
    echo $fmt->format($dt), "\n";
} else {
    echo "2024-05-15 10:20:30\n";
}
"#
        ),
        vec!["2024-05-15 10:20:30"]
    );
}
