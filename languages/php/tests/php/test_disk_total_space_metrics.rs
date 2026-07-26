
crate::php_cases! {
    disk_total_space_basic => {
        r#"<?php
$space = disk_total_space("/");
echo is_float($space) || is_int($space) ? "ok" : "fail";
"#,
        ["ok"]
    };
}
