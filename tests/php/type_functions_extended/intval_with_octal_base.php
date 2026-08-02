<?php
// vybe-test: php/type_functions_extended/intval_with_octal_base
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

$n = intval('777', 8);
echo $n;
