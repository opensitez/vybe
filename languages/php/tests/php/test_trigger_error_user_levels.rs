use super::helpers::run_prints;

crate::php_cases! {
    trigger_error_levels => {
        r#"<?php
set_error_handler(function($errno, $errstr) {
    echo "$errno:$errstr|";
    return true;
});

trigger_error("warn", E_USER_WARNING);
trigger_error("notice", E_USER_NOTICE);
trigger_error("deprecated", E_USER_DEPRECATED);
"#,
        ["512:warn|1024:notice|16384:deprecated|"]
    };
}
