<?php
// vybe-test: php/date_builtins/idate_access_patterns
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$ts = mktime(0, 0, 0, 4, 25, 2024);
echo idate('Y', $ts);
echo idate('m', $ts);
echo idate('d', $ts);
echo idate('w', $ts);
