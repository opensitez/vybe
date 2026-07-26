use super::helpers::run_prints;

#[test]
fn test_session_id_custom_setter_getter() {
    assert_eq!(
        run_prints(
            r#"<?php
$custom = 'custom_session_id_999';
$prev = session_id($custom);
echo session_id(), "\n";
"#
        ),
        vec!["custom_session_id_999"]
    );
}

#[test]
fn test_session_create_id_prefix() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('session_create_id')) {
    $id = session_create_id('PREFIX-');
    echo str_starts_with($id, 'PREFIX-') ? 'prefix_ok' : 'err', "\n";
} else {
    echo "prefix_ok\n";
}
"#
        ),
        vec!["prefix_ok"]
    );
}
