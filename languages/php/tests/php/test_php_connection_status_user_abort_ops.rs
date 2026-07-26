use super::helpers::run_prints;

#[test]
fn test_connection_status_normal() {
    assert_eq!(
        run_prints(
            r#"<?php
$st = connection_status();
echo ($st === CONNECTION_NORMAL || $st === 0) ? 'normal' : 'other', "\n";
"#
        ),
        vec!["normal"]
    );
}

#[test]
fn test_connection_aborted_false_cli() {
    assert_eq!(
        run_prints(
            r#"<?php
echo connection_aborted() === 0 ? 'not_aborted' : 'aborted', "\n";
"#
        ),
        vec!["not_aborted"]
    );
}

#[test]
fn test_ignore_user_abort_toggle() {
    assert_eq!(
        run_prints(
            r#"<?php
$prev = ignore_user_abort(true);
$now = ignore_user_abort(false);
echo ($now === 1 || $now === true) ? 'toggled' : 'err', "\n";
"#
        ),
        vec!["toggled"]
    );
}
