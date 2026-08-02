<?php
// vybe-test: php/type_juggling/coercion_concat_converts_to_string
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$n = 42;
$f = 3.14;
$b = true;
echo $n . "," . $f . "," . $b;
echo "\n";
echo null . "null";
