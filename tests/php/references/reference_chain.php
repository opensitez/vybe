<?php
// vybe-test: php/references/reference_chain
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$a = 1;
$b = &$a;
$c = &$b;
$c = 99;
echo $a;
