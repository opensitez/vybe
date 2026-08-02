<?php
// vybe-test: php/references/pass_by_reference_swap
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

function swap(&$a, &$b) { $tmp = $a; $a = $b; $b = $tmp; }
$x = "hello"; $y = "world";
swap($x, $y);
echo $x . " " . $y;
