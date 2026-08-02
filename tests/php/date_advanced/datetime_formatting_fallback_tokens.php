<?php
// vybe-test: php/date_advanced/datetime_formatting_fallback_tokens
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable('2024-02-29 23:59:59', new DateTimeZone('UTC'));
echo $dt->format('c');
echo '|' . $dt->format('Y-m-d');
echo '|' . $dt->format('l');
echo '|' . $dt->format('u');
