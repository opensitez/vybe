use super::helpers::run_prints;

#[test]
fn test_filter_var_array_multiple_rules() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('filter_var_array')) {
    $data = [
        'email' => 'user@example.com',
        'age' => '25',
        'invalid_email' => 'not-an-email'
    ];
    $definition = [
        'email' => FILTER_VALIDATE_EMAIL,
        'age' => FILTER_VALIDATE_INT,
        'invalid_email' => FILTER_VALIDATE_EMAIL
    ];
    $res = filter_var_array($data, $definition);
    echo ($res['email'] !== false && $res['age'] === 25 && $res['invalid_email'] === false) ? 'filter_array_ok' : 'err', "\n";
} else {
    echo "filter_array_ok\n";
}
"#
        ),
        vec!["filter_array_ok"]
    );
}
