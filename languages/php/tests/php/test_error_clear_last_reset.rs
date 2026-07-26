
crate::php_cases! {
    error_clear_last_basic => {
        r#"<?php
@trigger_error("test", E_USER_WARNING);
error_clear_last();
$err = error_get_last();
echo is_null($err) ? "null" : "not";
"#,
        ["null"]
    };
}
