<?php
// vybe-test: php/date_builtins/strtotime_parse_iso
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$ts = strtotime('2024-06-15');
echo $ts !== false ? 'parsed' : 'failed';
echo date('Y', $ts);
