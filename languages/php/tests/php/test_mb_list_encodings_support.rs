use super::helpers::run_prints;

crate::php_cases! {
    mb_list_encodings_basic => {
        r#"<?php
$list = mb_list_encodings();
echo in_array("UTF-8", $list) ? "found" : "missing";
"#,
        ["found"]
    };
}
