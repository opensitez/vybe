<?php
// vybe-test: php/scoping/dynamic_variable_reference_and_functions
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$name = 'target'; $fn = function() use ($name) { return $name; }; echo $fn();
$var = 'name'; echo '|' . $$var;
