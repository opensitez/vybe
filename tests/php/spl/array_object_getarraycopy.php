<?php
// vybe-test: php/spl/array_object_getarraycopy
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$ao = new ArrayObject([3, 1, 4, 1, 5]);
$copy = $ao->getArrayCopy();
sort($copy);
echo implode(',', $copy);
