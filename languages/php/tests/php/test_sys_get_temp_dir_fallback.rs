
crate::php_cases! {
    sys_get_temp_dir_basic => {
        r#"<?php
$dir = sys_get_temp_dir();
echo is_string($dir) && strlen($dir) > 0 ? "ok" : "fail";
"#,
        ["ok"]
    };
}
