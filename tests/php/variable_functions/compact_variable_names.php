<?php
// vybe-test: php/variable_functions/compact_variable_names
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$city    = 'Paris';
$country = 'France';
$pop     = 2161000;
$result = compact('city', 'country', 'pop');
echo $result['city'];
echo $result['country'];
