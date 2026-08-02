<?php
// vybe-test: php/date_builtins/date_isodate_parse_and_serialize
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable('2024-11-15T08:00:00+00:00');
echo $dt->format(DateTimeInterface::ATOM);
echo ':' . $dt->format('U');
