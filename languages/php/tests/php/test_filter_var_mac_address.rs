use super::helpers::run_prints;

crate::php_cases! {
    filter_var_mac_address_valid => {
        r#"<?php
echo filter_var("00:1A:2B:3C:4D:5E", FILTER_VALIDATE_MAC) ?: "fail";
"#,
        ["00:1A:2B:3C:4D:5E"]
    };

    filter_var_mac_address_invalid => {
        r#"<?php
echo filter_var("00:1A:2B:3C:4D:5Z", FILTER_VALIDATE_MAC) === false ? "fail" : "pass";
"#,
        ["fail"]
    };
}
