use super::helpers::run_prints;

crate::php_cases! {
    disk_free_space_basic => {
        r#"<?php
$space = disk_free_space("/");
echo is_float($space) || is_int($space) ? "ok" : "fail";
"#,
        ["ok"]
    };
}
