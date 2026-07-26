use super::helpers::run_prints;

#[test]
fn test_parse_url_malformed_returns_false() {
    assert_eq!(
        run_prints(
            r#"<?php
$res = @parse_url('http://:80');
echo ($res === false || is_array($res)) ? 'parse_url_handled' : 'err', "\n";
"#
        ),
        vec!["parse_url_handled"]
    );
}

#[test]
fn test_parse_url_scheme_relative() {
    assert_eq!(
        run_prints(
            r#"<?php
$res = parse_url('//cdn.example.com/app.js');
echo $res['host'] . ':' . $res['path'], "\n";
"#
        ),
        vec!["cdn.example.com:/app.js"]
    );
}
