<?php
// vybe-test: php/php_constants/uppercase_true_false_null
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$t = TRUE;
$f = FALSE;
$n = NULL;
echo $t ? 'yes' : 'no';
echo $f ? 'yes' : 'no';
echo is_null($n) ? 'null' : 'not null';
