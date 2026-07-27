crate::php_cases! {
    set_exception_handler_returns_previous => {
        r#"<?php
set_exception_handler(function($e) { echo "A"; });
$old = set_exception_handler(function($e) { echo "B"; });

echo is_callable($old) ? "callable" : "not";
"#,
        ["callable"]
    };
}
