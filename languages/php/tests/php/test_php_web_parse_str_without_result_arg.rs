use super::helpers::run_prints;

#[test]
fn test_parse_str_output_array_parameter() {
    assert_eq!(
        run_prints(
            r#"<?php
$queryString = "a=10&b[]=x&b[]=y&c[name]=Alice";
parse_str($queryString, $result);
echo $result['a'] . '|' . implode(',', $result['b']) . '|' . $result['c']['name'], "\n";
"#
        ),
        vec!["10|x,y|Alice"]
    );
}
