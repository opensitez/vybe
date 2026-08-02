<?php
// vybe-test: php/date_advanced/datetime_create_from_interface
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$mutable = new DateTime('2024-03-15');
$immutable = DateTimeImmutable::createFromMutable($mutable);
echo $immutable->format('Y-m-d');
echo ($immutable instanceof DateTimeImmutable) ? ':immutable' : ':not immutable';
