<?php
// vybe-test: php/date_advanced/datetime_timestamp
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable('2024-01-01 00:00:00', new DateTimeZone('UTC'));
$ts = $dt->getTimestamp();
echo $ts > 0 ? 'positive ts' : 'non-positive';
$back = (new DateTimeImmutable())->setTimestamp($ts);
echo ':' . $back->format('Y');
