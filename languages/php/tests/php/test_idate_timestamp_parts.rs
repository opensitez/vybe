use super::helpers::run_prints;

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
}
