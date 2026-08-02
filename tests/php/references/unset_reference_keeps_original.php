<?php
// vybe-test: php/references/unset_reference_keeps_original
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$x = 1;
$y = &$x;
unset($y);
$x = 2;
echo $x;
