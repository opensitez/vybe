<?php
// vybe-test: php/spl/array_object_append
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$ao = new ArrayObject([1, 2, 3]);
$ao->append(4);
$ao->append(5);
echo $ao->count();
