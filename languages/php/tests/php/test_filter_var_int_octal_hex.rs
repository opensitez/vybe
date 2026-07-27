crate::php_cases! {
    filter_var_int_basic => {
        r#"<?php
echo filter_var("123", FILTER_VALIDATE_INT);
"#,
        ["123"]
    };

    filter_var_int_hex => {
        r#"<?php
echo filter_var("0xFF", FILTER_VALIDATE_INT, FILTER_FLAG_ALLOW_HEX);
"#,
        ["255"]
    };

    filter_var_int_octal => {
        r#"<?php
echo filter_var("0123", FILTER_VALIDATE_INT, FILTER_FLAG_ALLOW_OCTAL);
"#,
        ["83"]
    };
}
