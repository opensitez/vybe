use super::helpers::run_prints;

#[test]
fn test_filter_input_array_returns_null_or_array() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('filter_input_array')) {
    $res = filter_input_array(INPUT_GET, ['id' => FILTER_VALIDATE_INT]);
    echo ($res === null || is_array($res) || $res === false) ? 'filter_input_array_ok' : 'err', "\n";
} else {
    echo "filter_input_array_ok\n";
}
"#
        ),
        vec!["filter_input_array_ok"]
    );
}
