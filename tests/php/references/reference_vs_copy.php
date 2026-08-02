<?php
// vybe-test: php/references/reference_vs_copy
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$a = [1, 2, 3];
$b = $a;     // copy
$c = &$a;    // reference
$b[] = 99;
$c[] = 88;
echo count($a) . ',' . count($b);
