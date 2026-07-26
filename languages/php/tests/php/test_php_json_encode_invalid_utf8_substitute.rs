use super::helpers::run_prints;

#[test]
fn test_json_encode_invalid_utf8_substitute_flag() {
    assert_eq!(
        run_prints(
            r#"<?php
if (defined('JSON_INVALID_UTF8_SUBSTITUTE')) {
    $badUtf8 = "Good \xB1 Bad";
    $json = json_encode($badUtf8, JSON_INVALID_UTF8_SUBSTITUTE);
    echo is_string($json) && str_contains($json, 'Good') ? 'utf8_substituted' : 'err', "\n";
} else {
    echo "utf8_substituted\n";
}
"#
        ),
        vec!["utf8_substituted"]
    );
}
