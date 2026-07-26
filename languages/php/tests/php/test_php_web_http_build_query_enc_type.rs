use super::helpers::run_prints;

#[test]
fn test_http_build_query_rfc3986_encoding() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['name' => 'John Doe', 'symbol' => 'foo+bar'];
echo http_build_query($data, '', '&', PHP_QUERY_RFC3986), "\n";
"#
        ),
        vec!["name=John%20Doe&symbol=foo%2Bbar"]
    );
}

#[test]
fn test_http_build_query_rfc1738_encoding() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['name' => 'John Doe'];
echo http_build_query($data, '', '&', PHP_QUERY_RFC1738), "\n";
"#
        ),
        vec!["name=John+Doe"]
    );
}
