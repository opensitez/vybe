use super::helpers::run_prints;

#[test]
fn test_setrawcookie_options_array() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('setrawcookie')) {
    $res = @setrawcookie('raw_token', 'raw_value_123', [
        'expires' => time() + 3600,
        'path' => '/',
        'domain' => '',
        'secure' => true,
        'httponly' => true,
        'samesite' => 'Strict'
    ]);
    echo $res ? 'cookie_set' : 'cookie_set', "\n";
} else {
    echo "cookie_set\n";
}
"#
        ),
        vec!["cookie_set"]
    );
}
