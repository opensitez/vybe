use super::helpers::run_prints;

crate::php_cases! {
    mb_str_split_basic => {
        r#"<?php
$str = "äöü";
$arr = mb_str_split($str, 1, "UTF-8");
echo implode('|', $arr);
"#,
        ["ä|ö|ü"]
    };

    mb_str_split_chunking => {
        r#"<?php
$str = "abcdef";
$arr = mb_str_split($str, 2);
echo implode('|', $arr);
"#,
        ["ab|cd|ef"]
    };
}
