crate::php_cases! {
    debug_print_backtrace_limit => {
        r#"<?php
function a() { b(); }
function b() { c(); }
function c() {
    ob_start();
    debug_print_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS, 2);
    $trace = ob_get_clean();
    echo substr_count($trace, '#') === 2 ? "ok" : "fail";
}
a();
"#,
        ["ok"]
    };
}
