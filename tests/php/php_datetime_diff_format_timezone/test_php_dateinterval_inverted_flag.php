<?php
// vybe-test: php/php_datetime_diff_format_timezone/test_php_dateinterval_inverted_flag
// origin: languages/php/tests/php/test_php_datetime_diff_format_timezone.rs
// vybe-test-mode: compile

$d1 = new DateTimeImmutable("2024-05-12");
$d2 = new DateTimeImmutable("2024-01-01");
$diff = $d1->diff($d2);
echo $diff->invert ? "INVERTED_PAST" : "FUTURE";
