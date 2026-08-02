<?php
// vybe-test: php/intl/message_formatter_basic
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('MessageFormatter')) { echo 'skipped'; return; }
$fmt = new MessageFormatter('en_US', 'Hello, {0}!');
echo $fmt->format(['World']);
