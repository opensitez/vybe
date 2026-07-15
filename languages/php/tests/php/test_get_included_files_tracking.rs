use super::helpers::run_prints;

crate::php_cases! {
    get_included_files_basic => {
        r#"<?php
$files = get_included_files();
echo is_array($files) && count($files) >= 1 ? "ok" : "fail";
"#,
        ["ok"]
    };

    get_required_files_alias => {
        r#"<?php
$files = get_required_files();
echo is_array($files) && count($files) >= 1 ? "ok" : "fail";
"#,
        ["ok"]
    };
}
