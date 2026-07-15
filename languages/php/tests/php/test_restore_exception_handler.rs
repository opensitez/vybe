use super::helpers::run_prints;

crate::php_cases! {
    restore_exception_handler_basic => {
        r#"<?php
set_exception_handler(function($e) { echo "A"; });
set_exception_handler(function($e) { echo "B"; });
restore_exception_handler();

$old = set_exception_handler(function($e) { echo "C"; });
// We just check if it's callable without throwing
echo "ok";
"#,
        ["ok"]
    };
}
