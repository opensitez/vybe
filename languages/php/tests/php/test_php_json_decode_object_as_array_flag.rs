use super::helpers::run_prints;

#[test]
fn test_json_decode_object_as_array_flag() {
    assert_eq!(
        run_prints(
            r#"<?php
$json = '{"name":"Alice","items":[1,2,3]}';
$decoded = json_decode($json, false, 512, JSON_OBJECT_AS_ARRAY);
echo is_array($decoded) && $decoded['name'] === 'Alice' ? 'array_decoded' : 'err', "\n";
"#
        ),
        vec!["array_decoded"]
    );
}
