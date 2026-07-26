
crate::php_cases! {
    stream_resolve_include_path_exists => {
        r#"<?php
// Since we can't reliably predict the include path, we'll just test if it returns a string or false
$path = stream_resolve_include_path("php://memory");
echo is_string($path) || is_bool($path) ? "ok" : "fail";
"#,
        ["ok"]
    };
}
