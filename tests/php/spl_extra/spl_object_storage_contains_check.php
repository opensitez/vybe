<?php
// vybe-test: php/spl_extra/spl_object_storage_contains_check
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
$obj1 = new stdClass(); $obj1->id = 1;
$obj2 = new stdClass(); $obj2->id = 2;
$store->attach($obj1);
echo $store->contains($obj1) ? 'yes' : 'no';
echo $store->contains($obj2) ? 'yes' : 'no';
$store->attach($obj2);
echo $store->contains($obj2) ? 'yes' : 'no';
