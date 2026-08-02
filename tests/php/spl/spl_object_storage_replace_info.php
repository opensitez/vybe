<?php
// vybe-test: php/spl/spl_object_storage_replace_info
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, 'first');
$store->rewind();
$store->setInfo('second');
echo $store->getInfo();
