<?php
// vybe-test: php/date_advanced/date_interval_components
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$i = new DateInterval('P3Y6M15DT12H30M45S');
echo $i->y . ':' . $i->m . ':' . $i->d;
echo ':' . $i->h . ':' . $i->i . ':' . $i->s;
