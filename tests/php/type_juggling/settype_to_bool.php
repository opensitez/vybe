<?php
// vybe-test: php/type_juggling/settype_to_bool
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$v = 0;
settype($v, 'bool');
var_dump($v);
$v2 = "hello";
settype($v2, 'boolean');
var_dump($v2);
