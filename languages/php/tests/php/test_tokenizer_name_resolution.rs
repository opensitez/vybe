crate::php_cases! {
    token_name_resolves_constants => {
        r#"<?php
echo token_name(T_CLASS) . '|' . token_name(T_FUNCTION) . '|' . token_name(T_STRING);
"#,
        ["T_CLASS|T_FUNCTION|T_STRING"]
    };

    token_name_unknown_token => {
        r#"<?php
// PHP returns "UNKNOWN" for invalid tokens
echo token_name(999999);
"#,
        ["UNKNOWN"]
    };
}
