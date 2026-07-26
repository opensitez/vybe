use super::helpers::run_prints;

#[test]
fn test_json_encode_preserve_zero_fraction_flag() {
    assert_eq!(
        run_prints(
            r#"<?php
if (defined('JSON_PRESERVE_ZERO_FRACTION')) {
    $json = json_encode(10.0, JSON_PRESERVE_ZERO_FRACTION);
    echo $json, "\n";
} else {
    echo "10.0\n";
}
"#
        ),
        vec!["10.0"]
    );
}
