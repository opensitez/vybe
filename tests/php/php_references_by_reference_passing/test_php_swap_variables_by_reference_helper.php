<?php
// vybe-test: php/php_references_by_reference_passing/test_php_swap_variables_by_reference_helper
// origin: languages/php/tests/php/test_php_references_by_reference_passing.rs
// vybe-test-mode: compile

function swap(&$a, &$b) {
    $tmp = $a;
    $a = $b;
    $b = $tmp;
}

$x = "first"; $y = "second";
swap($x, $y);
echo "$x $y";
