use super::helpers::run_prints;

crate::php_cases! {
    getrusage_basic => {
        r#"<?php
$usage = getrusage();
echo is_array($usage) && isset($usage['ru_utime.tv_usec']) ? "ok" : "fail";
"#,
        ["ok"]
    };

    getrusage_children => {
        r#"<?php
$usage = getrusage(1); // RUSAGE_CHILDREN
echo is_array($usage) ? "ok" : "fail";
"#,
        ["ok"]
    };
}
