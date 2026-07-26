use super::helpers::run_prints;

#[test]
fn test_array_uassoc_custom_key_comparison() {
    assert_eq!(
        run_prints(
            r#"<?php
$a1 = ['a' => 1, 'b' => 2];
$a2 = ['A' => 1, 'B' => 3];
$diff = array_uassoc($a1, $a2, 'strcasecmp');
echo count($diff) . ':' . implode(',', array_keys($diff)), "\n";
"#
        ),
        vec!["1:b"]
    );
}
