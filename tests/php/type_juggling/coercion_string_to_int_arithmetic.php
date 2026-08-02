<?php
// vybe-test: php/type_juggling/coercion_string_to_int_arithmetic
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$a = "5";
$b = $a + 3;
var_dump($b);       // int(8)
$c = "5.5" + 1;
var_dump($c);       // float(6.5)
$d = "5 apples" + 2;
var_dump($d);       // int(7)
$e = "apples" + 2;
var_dump($e);       // int(2)
