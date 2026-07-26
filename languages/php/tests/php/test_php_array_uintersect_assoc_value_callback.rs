use super::helpers::run_prints;

#[test]
fn test_array_uintersect_assoc_custom_value_comparison() {
    assert_eq!(
        run_prints(
            r#"<?php
$a1 = ['x' => '10', 'y' => '20'];
$a2 = ['x' => 10, 'y' => 30];
$intersection = array_uintersect_assoc($a1, $a2, function($v1, $v2) {
    return (int)$v1 <=> (int)$v2;
});
echo count($intersection) . ':' . implode(',', array_keys($intersection)), "\n";
"#
        ),
        vec!["1:x"]
    );
}
