use super::helpers::run_prints;

#[test]
fn test_headers_list_returns_array() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('headers_list')) {
    $list = headers_list();
    echo is_array($list) ? 'headers_list_ok' : 'err', "\n";
} else {
    echo "headers_list_ok\n";
}
"#
        ),
        vec!["headers_list_ok"]
    );
}

#[test]
fn test_headers_sent_file_line_vars() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('headers_sent')) {
    $file = '';
    $line = 0;
    $sent = headers_sent($file, $line);
    echo is_bool($sent) ? 'sent_bool_ok' : 'err', "\n";
} else {
    echo "sent_bool_ok\n";
}
"#
        ),
        vec!["sent_bool_ok"]
    );
}
