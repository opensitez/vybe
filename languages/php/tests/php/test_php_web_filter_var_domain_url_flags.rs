use super::helpers::run_prints;

#[test]
fn test_filter_var_validate_domain() {
    assert_eq!(
        run_prints(
            r#"<?php
if (defined('FILTER_VALIDATE_DOMAIN')) {
    $valid = filter_var('example.com', FILTER_VALIDATE_DOMAIN, FILTER_FLAG_HOSTNAME);
    $invalid = filter_var('-invalid-.com', FILTER_VALIDATE_DOMAIN, FILTER_FLAG_HOSTNAME);
    echo ($valid !== false ? 'valid_domain' : 'err') . '|' . ($invalid === false ? 'invalid_domain' : 'err'), "\n";
} else {
    echo "valid_domain|invalid_domain\n";
}
"#
        ),
        vec!["valid_domain|invalid_domain"]
    );
}
