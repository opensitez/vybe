<?php
// vybe-test: php/intl/message_formatter_plural
// origin: languages/php/tests/php/test_intl.rs
// vybe-test-mode: compile

if (!class_exists('MessageFormatter')) { echo 'skipped'; return; }
$pattern = '{0, plural, =0{no items} one{# item} other{# items}}';
$fmt = new MessageFormatter('en_US', $pattern);
echo $fmt->format([0]) . ':';
echo $fmt->format([1]) . ':';
echo $fmt->format([5]);
