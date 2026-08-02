<?php
// vybe-test: php/php_datetime_diff_format_timezone/test_php_datetime_microtime_precision
// origin: languages/php/tests/php/test_php_datetime_diff_format_timezone.rs
// vybe-test-mode: compile

$dt = DateTimeImmutable::createFromFormat("U.u", sprintf("%.6f", microtime(true)));
echo $dt->format("Y-m-d H:i:s.u");
