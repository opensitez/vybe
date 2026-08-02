<?php
// vybe-test: php/date_advanced/datetime_modify_with_relative_words
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable('2024-01-15 10:00:00');
echo $dt->modify('first day of next month')->format('Y-m-d');
echo '|' . $dt->modify('midnight')->format('H:i:s');
