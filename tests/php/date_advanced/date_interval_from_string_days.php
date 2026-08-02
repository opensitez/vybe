<?php
// vybe-test: php/date_advanced/date_interval_from_string_days
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$i = DateInterval::createFromDateString('3 weeks + 2 days');
echo $i->y . ':' . $i->m . ':' . $i->d;
