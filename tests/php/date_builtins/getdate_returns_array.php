<?php
// vybe-test: php/date_builtins/getdate_returns_array
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

date_default_timezone_set('UTC');
$ts = mktime(10, 30, 0, 7, 4, 2024);
$info = getdate($ts);
echo $info['year'];
echo $info['mon'];
echo $info['mday'];
