<?php
// vybe-test: php/type_juggling/strict_vs_loose
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$a = 0; $b = false;
echo ($a == $b)  ? "loose equal\n" : "loose not equal\n";
echo ($a === $b) ? "strict equal\n" : "strict not equal\n";
$c = "1"; $d = 1;
echo ($c == $d)  ? "loose equal\n" : "loose not equal\n";
echo ($c === $d) ? "strict equal\n" : "strict not equal\n";
