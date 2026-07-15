use super::helpers::run_prints;

crate::php_cases! {
    strtotime_relative_keywords => {
        r#"<?php
date_default_timezone_set('UTC');
$base = strtotime('2020-01-15 10:00:00');
echo date('Y-m-d H:i:s', strtotime('+1 day', $base)) . "|";
echo date('Y-m-d H:i:s', strtotime('next monday', $base));
"#,
        ["2020-01-16 10:00:00|2020-01-20 00:00:00"]
    };
}
