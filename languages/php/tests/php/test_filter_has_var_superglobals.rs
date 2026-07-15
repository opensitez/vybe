use super::helpers::run_prints;

crate::php_cases! {
    filter_has_var_basic => {
        r#"<?php
// We test if the function exists and can handle empty superglobals gracefully
echo filter_has_var(INPUT_GET, 'test') ? "yes" : "no";
"#,
        ["no"]
    };
}
