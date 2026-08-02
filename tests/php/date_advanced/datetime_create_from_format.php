<?php
// vybe-test: php/date_advanced/datetime_create_from_format
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$dt = DateTime::createFromFormat('d/m/Y H:i', '15/06/2024 14:30');
echo $dt->format('Y-m-d H:i');
$dt2 = DateTimeImmutable::createFromFormat('U', '1718438400');
echo ':' . ($dt2 !== false ? 'ok' : 'fail');
