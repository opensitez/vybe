<?php
// vybe-test: php/date_advanced/datetime_interval_roundtrip_iso8601
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$interval = new DateInterval('P1Y2M3DT4H5M6S');
$spec = $interval->format('P%yY%mM%dDT%hH%iM%sS');
echo $spec;
echo '|' . (DateInterval::createFromDateString('2 weeks') !== false ? 'from_string' : 'from_string_failed');
echo '|' . (new DateTimeImmutable('2024-01-01'))->add($interval)->format('Y');
