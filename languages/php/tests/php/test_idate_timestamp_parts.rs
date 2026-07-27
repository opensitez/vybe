crate::php_cases! {
    idate_various_parts => {
        r#"<?php
date_default_timezone_set('UTC');
$timestamp = mktime(14, 30, 45, 10, 24, 2020);
echo idate('Y', $timestamp) . "|";
echo idate('m', $timestamp) . "|";
echo idate('d', $timestamp) . "|";
echo idate('H', $timestamp);
"#,
        ["2020|10|24|14"]
    };

    idate_week_of_year => {
        r#"<?php
date_default_timezone_set('UTC');
$timestamp = mktime(0, 0, 0, 1, 1, 2024);
echo idate('W', $timestamp);
"#,
        ["1"]
    };

    idate_day_of_week => {
        r#"<?php
date_default_timezone_set('UTC');
$timestamp = mktime(0, 0, 0, 10, 24, 2020);
echo idate('w', $timestamp) . "|" . idate('N', $timestamp);
"#,
        ["6|6"]
    };

    idate_month_day_names_runtime => {
        r#"<?php
date_default_timezone_set('UTC');
$timestamp = mktime(0, 0, 0, 12, 31, 2023);
echo idate('L', $timestamp) . "|" . idate('t', $timestamp);
"#,
        ["1|31"]
    };
}
