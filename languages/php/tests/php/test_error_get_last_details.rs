use super::helpers::run_prints;

crate::php_cases! {
    error_get_last_keys => {
        r#"<?php
@trigger_error("custom msg", E_USER_NOTICE);
$err = error_get_last();
if ($err) {
    echo $err['type'] === E_USER_NOTICE ? "notice|" : "fail|";
    echo $err['message'] === "custom msg" ? "msg" : "fail";
}
"#,
        ["notice|msg"]
    };
}
