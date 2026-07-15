use super::helpers::run_prints;

crate::php_cases! {
    stream_set_blocking_mode => {
        r#"<?php
$fp = fopen("php://temp", "r+");
echo stream_set_blocking($fp, false) ? "ok|" : "fail|";
$meta = stream_get_meta_data($fp);
echo $meta['blocked'] ? "blocked" : "unblocked";
"#,
        ["ok|unblocked"]
    };
}
