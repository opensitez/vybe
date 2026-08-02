<?php
// vybe-test: php/references/is_same_reference
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$a = [1, 2, 3];
$b = &$a;
$b[] = 4;
echo count($a);
