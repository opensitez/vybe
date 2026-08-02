<?php
// vybe-test: php/spl/spl_object_storage_info_api
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, ['tag' => 'alpha']);
echo $store->contains($obj) ? 'yes' : 'no';
echo $store->getInfo()['tag'];
echo $store->getHash($obj);
