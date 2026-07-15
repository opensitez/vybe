use super::helpers::run_prints;

crate::php_cases! {
    filter_var_domain_valid => {
        r#"<?php
echo filter_var("example.com", FILTER_VALIDATE_DOMAIN) ?: "fail";
"#,
        ["example.com"]
    };

    filter_var_domain_hostname_flag => {
        r#"<?php
echo filter_var("http://example.com", FILTER_VALIDATE_DOMAIN, FILTER_FLAG_HOSTNAME) === false ? "fail" : "pass";
"#,
        ["fail"] // http:// is not a hostname
    };
}
