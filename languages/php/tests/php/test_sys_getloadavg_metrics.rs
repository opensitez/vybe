crate::php_cases! {
    sys_getloadavg_basic => {
        r#"<?php
$load = sys_getloadavg();
echo is_array($load) && count($load) === 3 ? "ok" : "fail";
"#,
        ["ok"]
    };
}
