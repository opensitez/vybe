<?php
// vybe-test: php/date_advanced/date_interval_format
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$i = new DateInterval('P1Y2M3DT4H5M6S');
echo $i->format('%Y years, %M months, %D days');
echo ':' . $i->format('%H:%I:%S');
