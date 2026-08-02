<?php
// vybe-test: php/references/reference_assignment_basic
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$a = 10;
$b = &$a;
$b = 20;
echo $a;
