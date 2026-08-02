<?php
// vybe-test: php/type_functions_extended/intval_with_hex_base
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

$n = intval('1F', 16);
echo $n;
