<?php
// vybe-test: php/type_juggling/settype_to_int
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$v = "42abc";
settype($v, 'integer');
var_dump($v);
