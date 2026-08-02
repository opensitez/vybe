<?php
// vybe-test: php/type_juggling/settype_to_string
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$v = 123;
settype($v, 'string');
var_dump($v);
