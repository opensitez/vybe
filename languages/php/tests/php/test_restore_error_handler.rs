
crate::php_cases! {
    restore_error_handler_basic => {
        r#"<?php
set_error_handler(function() { echo "A"; });
set_error_handler(function() { echo "B"; });
restore_error_handler();

@trigger_error("msg", E_USER_NOTICE);
"#,
        ["A"]
    };
}
