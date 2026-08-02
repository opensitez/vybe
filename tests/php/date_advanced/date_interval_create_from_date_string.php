<?php
// vybe-test: php/date_advanced/date_interval_create_from_date_string
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$i = DateInterval::createFromDateString('2 weeks + 3 days');
echo $i->days >= 0 ? 'created' : 'failed';
