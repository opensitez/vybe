use super::helpers::run_prints;

crate::php_cases! {
    memory_get_usage_basic => {
        r#"<?php
$mem = memory_get_usage();
echo is_int($mem) && $mem > 0 ? "ok" : "fail";
"#,
        ["ok"]
    };

    memory_get_usage_real => {
        r#"<?php
$mem = memory_get_usage(true);
echo is_int($mem) && $mem > 0 ? "ok" : "fail";
"#,
        ["ok"]
    };
}
