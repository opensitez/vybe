<?php
// vybe-test: php/type_juggling/settype_to_array
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$v = 42;
settype($v, 'array');
var_dump($v);
