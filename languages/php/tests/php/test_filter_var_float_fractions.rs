crate::php_cases! {
    filter_var_float_basic => {
        r#"<?php
echo filter_var("123.45", FILTER_VALIDATE_FLOAT) ?: "fail";
"#,
        ["123.45"]
    };

    filter_var_float_fractions => {
        r#"<?php
echo filter_var("1,234.56", FILTER_VALIDATE_FLOAT, FILTER_FLAG_ALLOW_THOUSAND) ?: "fail";
"#,
        ["1234.56"]
    };
}
