<?php
// vybe-test: php/spl/spl_object_storage_detach
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
$a = new stdClass();
$b = new stdClass();
$store->attach($a);
$store->attach($b);
$store->detach($a);
echo $store->count();
