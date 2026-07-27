//! `parse_url`, `http_build_query`, `rawurlencode`, and query parsing.

crate::php_cases! {
    parse_url_extracts_scheme_host_port => {
        r#"<?php
$p = parse_url('https://example.com:8080/path');
echo $p['scheme'] . ':' . $p['host'] . ':' . $p['port'];
"#,
        ["https:example.com:8080"]
    };

    parse_url_php_url_host_component => {
        r#"<?php
echo parse_url('https://sub.example.com/x', PHP_URL_HOST);
"#,
        ["sub.example.com"]
    };

    parse_url_php_url_path_component => {
        r#"<?php
echo parse_url('https://example.com/foo/bar?q=1', PHP_URL_PATH);
"#,
        ["/foo/bar"]
    };

    parse_url_php_url_query_component => {
        r#"<?php
echo parse_url('https://example.com/?q=hello&p=2', PHP_URL_QUERY);
"#,
        ["q=hello&p=2"]
    };

    parse_url_php_url_fragment_component => {
        r#"<?php
echo parse_url('https://example.com/#top', PHP_URL_FRAGMENT);
"#,
        ["top"]
    };

    parse_url_userinfo_username => {
        r#"<?php
echo parse_url('http://user:pass@host/', PHP_URL_USER);
"#,
        ["user"]
    };

    http_build_query_assoc_array => {
        r#"<?php
echo http_build_query(['a' => 1, 'b' => 2]);
"#,
        ["a=1&b=2"]
    };

    http_build_query_nested_array_brackets => {
        r#"<?php
echo http_build_query(['user' => ['name' => 'ada', 'id' => 3]]);
"#,
        ["user%5Bname%5D=ada&user%5Bid%5D=3"]
    };

    http_build_query_custom_arg_separator => {
        r#"<?php
echo http_build_query(['x' => 1, 'y' => 2], '', '|');
"#,
        ["x=1|y=2"]
    };

    http_build_query_encodes_space_as_plus => {
        r#"<?php
echo http_build_query(['q' => 'a b']);
"#,
        ["q=a+b"]
    };

    rawurlencode_encodes_slash => {
        r#"<?php
echo rawurlencode('/a b');
"#,
        ["%2Fa%20b"]
    };

    rawurldecode_reverses_percent_encoding => {
        r#"<?php
echo rawurldecode('%2F%20');
"#,
        ["/ "]
    };

    urlencode_encodes_space_as_plus => {
        r#"<?php
echo urlencode('a b');
"#,
        ["a+b"]
    };

    urldecode_decodes_plus_as_space => {
        r#"<?php
echo urldecode('a+b');
"#,
        ["a b"]
    };

    parse_str_populates_variables => {
        r#"<?php
parse_str('foo=bar&n=9', $out);
echo $out['foo'] . ':' . $out['n'];
"#,
        ["bar:9"]
    };

    base64_encode_decode_roundtrip => {
        r#"<?php
echo base64_decode(base64_encode('vybe'));
"#,
        ["vybe"]
    };

    rfc3986_path_join_pattern => {
        r#"<?php
function join_path(string $base, string $path): string {
    return rtrim($base, '/') . '/' . ltrim($path, '/');
}
echo join_path('https://example.com/api/', '/users');
"#,
        ["https://example.com/api/users"]
    };

    parse_url_relative_path_only => {
        r#"<?php
$p = parse_url('/only/path');
echo $p['path'];
"#,
        ["/only/path"]
    };

    http_build_query_bool_false_omitted_or_zero => {
        r#"<?php
echo http_build_query(['on' => true, 'off' => false]);
"#,
        ["on=1&off=0"]
    };

    parse_url_missing_scheme_returns_no_scheme_key => {
        r#"<?php
echo array_key_exists('scheme', parse_url('//example.com')) ? 'yes' : 'no';
"#,
        ["no"]
    };

    http_build_query_numeric_prefix => {
        r#"<?php
echo http_build_query(['first', 'second'], 'item_');
"#,
        ["item_0=first&item_1=second"]
    };

    base64_decode_strict_mode => {
        r#"<?php
$res = base64_decode('invalid-base64!', true);
echo $res === false || is_string($res) ? 'strict_decode_ok' : 'err';
"#,
        ["strict_decode_ok"]
    };
}
