
crate::php_cases! {
    mb_check_encoding_valid => {
        r#"<?php
$str = "äöü";
echo mb_check_encoding($str, "UTF-8") ? "valid" : "invalid";
"#,
        ["valid"]
    };

    mb_check_encoding_invalid => {
        r#"<?php
$str = "\xc3\x28"; // Invalid UTF-8 sequence
echo mb_check_encoding($str, "UTF-8") ? "valid" : "invalid";
"#,
        ["invalid"]
    };
}
