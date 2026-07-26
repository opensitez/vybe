
crate::php_cases! {
    memory_get_peak_usage_basic => {
        r#"<?php
$peak = memory_get_peak_usage();
echo is_int($peak) && $peak > 0 ? "ok" : "fail";
"#,
        ["ok"]
    };

    memory_get_peak_usage_real => {
        r#"<?php
$peak = memory_get_peak_usage(true);
echo is_int($peak) && $peak > 0 ? "ok" : "fail";
"#,
        ["ok"]
    };
}
