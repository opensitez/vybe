<?php
// vybe-test: php/spl/spl_object_storage_basic
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
$a = new stdClass(); $a->name = 'Alice';
$b = new stdClass(); $b->name = 'Bob';
$store->attach($a, 'data-a');
$store->attach($b, 'data-b');
echo $store->count();
