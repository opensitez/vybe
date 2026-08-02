<?php
// vybe-test: php/spl_extra/spl_object_storage_attach_detach
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
$a = new stdClass();
$b = new stdClass();
$store->attach($a);
$store->attach($b);
echo $store->count();
$store->detach($a);
echo $store->count();
