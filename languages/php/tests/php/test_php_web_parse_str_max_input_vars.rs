use super::helpers::run_prints;

#[test]
fn test_parse_str_deeply_nested_arrays() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "a[b][c][d]=val";
parse_str($str, $res);
echo $res['a']['b']['c']['d'], "\n";
"#
        ),
        vec!["val"]
    );
}
