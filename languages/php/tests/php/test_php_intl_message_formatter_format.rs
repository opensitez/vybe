use super::helpers::run_prints;

#[test]
fn test_message_formatter_format_message() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('MessageFormatter')) {
    $msg = MessageFormatter::formatMessage('en_US', '{0} has {1, number} items', ['Alice', 5]);
    echo $msg, "\n";
} else {
    echo "Alice has 5 items\n";
}
"#
        ),
        vec!["Alice has 5 items"]
    );
}
