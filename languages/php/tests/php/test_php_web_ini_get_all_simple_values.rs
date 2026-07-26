use super::helpers::run_prints;

#[test]
fn test_ini_get_all_details_false_returns_strings() {
    assert_eq!(
        run_prints(
            r#"<?php
$all = ini_get_all(null, false);
echo is_array($all) && is_string($all['display_errors'] ?? '') ? 'simple_ini_ok' : 'err', "\n";
"#
        ),
        vec!["simple_ini_ok"]
    );
}
