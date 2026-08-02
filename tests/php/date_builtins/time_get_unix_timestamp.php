<?php
// vybe-test: php/date_builtins/time_get_unix_timestamp
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$t = time();
echo is_int($t) ? 'integer' : 'not integer';
echo $t > 1000000000 ? ':plausible' : ':implausible';
