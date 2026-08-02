<?php
// vybe-test: php/intl/message_formatter_date
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('MessageFormatter')) { echo 'skipped'; return; }
$fmt = new MessageFormatter('en_US', 'Date: {0, date, short}');
$result = $fmt->format([time()]);
echo strlen($result) > 0 ? 'date formatted' : 'empty';
