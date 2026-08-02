<?php
// vybe-test: php/spl/array_object_basic
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$ao = new ArrayObject(['x' => 1, 'y' => 2]);
echo $ao['x'];
$ao['z'] = 3;
echo $ao->count();
