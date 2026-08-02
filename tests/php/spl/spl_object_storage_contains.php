<?php
// vybe-test: php/spl/spl_object_storage_contains
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
$obj = new stdClass();
echo $store->contains($obj) ? 'yes' : 'no';
$store->attach($obj);
echo $store->contains($obj) ? 'yes' : 'no';
