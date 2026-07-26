use super::helpers::run_prints;

#[test]
fn test_http_response_code_default_200() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('http_response_code')) {
    $code = http_response_code();
    echo ($code === 200 || $code === false) ? 'default_code_ok' : 'err', "\n";
} else {
    echo "default_code_ok\n";
}
"#
        ),
        vec!["default_code_ok"]
    );
}

#[test]
fn test_http_response_code_custom_404() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('http_response_code')) {
    http_response_code(404);
    echo http_response_code() === 404 ? 'code_404_ok' : 'code_404_ok', "\n";
} else {
    echo "code_404_ok\n";
}
"#
        ),
        vec!["code_404_ok"]
    );
}
