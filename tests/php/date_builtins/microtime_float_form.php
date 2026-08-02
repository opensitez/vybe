<?php
// vybe-test: php/date_builtins/microtime_float_form
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$mt = microtime(true);
echo is_float($mt) ? 'float' : 'not float';
echo $mt > 1000000000.0 ? ':plausible' : ':implausible';
