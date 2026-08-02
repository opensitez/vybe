<?php
// vybe-test: php/spl/array_object_iterator_mode
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$ao = new ArrayObject(['a' => 1, 'b' => 2], ArrayObject::STD_PROP_LIST);
$ao->ksort();
$it = $ao->getIterator();
foreach ($it as $k => $v) { echo $k . $v; }
