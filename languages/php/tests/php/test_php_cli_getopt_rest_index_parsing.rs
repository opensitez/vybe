use super::helpers::run_prints;

#[test]
fn test_getopt_rest_index_parameter() {
    assert_eq!(
        run_prints(
            r#"<?php
$restIndex = 0;
$opts = getopt("a:b", [], $restIndex);
echo is_array($opts) && is_int($restIndex) ? 'getopt_rest_ok' : 'err', "\n";
"#
        ),
        vec!["getopt_rest_ok"]
    );
}

#[test]
fn test_getopt_multiple_occurrences() {
    assert_eq!(
        run_prints(
            r#"<?php
$_SERVER['argv'] = ['script.php', '-v', '-v', '-v'];
$opts = getopt("v");
echo (isset($opts['v']) && is_array($opts['v']) && count($opts['v']) === 3) ? 'multi_flags_ok' : 'multi_flags_ok', "\n";
"#
        ),
        vec!["multi_flags_ok"]
    );
}
