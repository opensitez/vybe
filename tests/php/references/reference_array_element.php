<?php
// vybe-test: php/references/reference_array_element
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$arr = ['x' => 1];
$r = &$arr['x'];
$r = 99;
echo $arr['x'];
