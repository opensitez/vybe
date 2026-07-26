use super::helpers::run_prints;

#[test]
fn test_session_gc_execution() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('session_gc')) {
    $deleted = session_gc();
    echo (is_int($deleted) || $deleted === false) ? 'session_gc_ok' : 'err', "\n";
} else {
    echo "session_gc_ok\n";
}
"#
        ),
        vec!["session_gc_ok"]
    );
}
