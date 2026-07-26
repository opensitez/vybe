
crate::php_cases! {
    filter_var_ipv4 => {
        r#"<?php
echo filter_var("192.168.1.1", FILTER_VALIDATE_IP, FILTER_FLAG_IPV4) ?: "fail";
"#,
        ["192.168.1.1"]
    };

    filter_var_ipv6 => {
        r#"<?php
echo filter_var("2001:0db8:85a3:0000:0000:8a2e:0370:7334", FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) ?: "fail";
"#,
        ["2001:0db8:85a3:0000:0000:8a2e:0370:7334"]
    };

    filter_var_ip_priv_range => {
        r#"<?php
echo filter_var("192.168.1.1", FILTER_VALIDATE_IP, FILTER_FLAG_NO_PRIV_RANGE) === false ? "filtered" : "pass";
"#,
        ["filtered"]
    };
}
