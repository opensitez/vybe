crate::php_cases! {
    set_error_handler_returns_false_bypasses => {
        r#"<?php
set_error_handler(function() {
    echo "caught|";
    return false; // Tells PHP to continue with normal error handler
});

@trigger_error("msg", E_USER_NOTICE);
echo "done";
"#,
        ["caught|done"]
    };
}
