crate::php_cases! {
    filter_var_email_valid => {
        r#"<?php
echo filter_var("test@example.com", FILTER_VALIDATE_EMAIL) ?: "fail";
"#,
        ["test@example.com"]
    };

    filter_var_email_unicode => {
        r#"<?php
echo filter_var("test@exämple.com", FILTER_VALIDATE_EMAIL, FILTER_FLAG_EMAIL_UNICODE) ?: "fail";
"#,
        ["test@exämple.com"]
    };
}
