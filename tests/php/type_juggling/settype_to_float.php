<?php
// vybe-test: php/type_juggling/settype_to_float
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$v = "3.99";
settype($v, 'float');
var_dump($v);
