
crate::php_cases! {
    stream_is_local_file => {
        r#"<?php
$fp = fopen("php://temp", "r");
echo stream_is_local($fp) ? "local" : "remote";
"#,
        ["local"]
    };

    stream_is_local_http => {
        r#"<?php
echo stream_is_local("http://example.com") ? "local" : "remote";
"#,
        ["remote"]
    };
}
