<?php
// vybe-test: php/intl/message_formatter_static
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('MessageFormatter')) { echo 'skipped'; return; }
$result = MessageFormatter::formatMessage('en_US', 'Value: {0, number}', [1234567.89]);
echo strlen($result) > 0 ? 'formatted' : 'empty';
