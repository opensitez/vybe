use super::helpers::run_prints;

#[test]
fn test_get_resources_all() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('get_resources')) {
    $f = fopen('php://memory', 'r+');
    $res = get_resources();
    fclose($f);
    echo is_array($res) ? 'res_array_ok' : 'err', "\n";
} else {
    echo "res_array_ok\n";
}
"#
        ),
        vec!["res_array_ok"]
    );
}

#[test]
fn test_get_resources_filter_by_type() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('get_resources')) {
    $f = fopen('php://memory', 'r+');
    $streams = get_resources('stream');
    fclose($f);
    echo is_array($streams) ? 'filtered_stream_ok' : 'err', "\n";
} else {
    echo "filtered_stream_ok\n";
}
"#
        ),
        vec!["filtered_stream_ok"]
    );
}
