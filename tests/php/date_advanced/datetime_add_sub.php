<?php
// vybe-test: php/date_advanced/datetime_add_sub
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable('2024-01-01');
$plus30 = $dt->add(new DateInterval('P30D'));
$minus7 = $dt->sub(new DateInterval('P7D'));
echo $plus30->format('Y-m-d');
echo ':' . $minus7->format('Y-m-d');
