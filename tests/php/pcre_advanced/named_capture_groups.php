<?php
// vybe-test: php/pcre_advanced/named_capture_groups
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$date = '2024-06-15';
preg_match('/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/', $date, $m);
echo $m['year'] . '-' . $m['month'] . '-' . $m['day'];
