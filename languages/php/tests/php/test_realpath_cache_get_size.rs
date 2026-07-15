use super::helpers::run_prints;

crate::php_cases! {
    realpath_cache_get_type => {
        r#"<?php
$cache = realpath_cache_get();
echo is_array($cache) ? "ok" : "fail";
"#,
        ["ok"]
    };

    realpath_cache_size_type => {
        r#"<?php
$size = realpath_cache_size();
echo is_int($size) ? "ok" : "fail";
"#,
        ["ok"]
    };
}
