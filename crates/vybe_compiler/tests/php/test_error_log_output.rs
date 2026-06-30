//! `error_log` — must route to stderr/logging, not PHP stdout.

crate::php_cases! {
    error_log_message_not_on_stdout => {
        r#"<?php
error_log('diagnostic');
echo 'visible';
"#,
        ["visible"]
    };

    error_log_does_not_break_following_echo => {
        r#"<?php
error_log('a');
error_log('b');
echo 'c';
"#,
        ["c"]
    };
}
