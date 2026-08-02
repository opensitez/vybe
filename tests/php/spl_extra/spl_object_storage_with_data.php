<?php
// vybe-test: php/spl_extra/spl_object_storage_with_data
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
$obj = new stdClass();
$store->attach($obj, ['role' => 'admin', 'active' => true]);
$store->rewind();
$data = $store->getInfo();
echo isset($data['role']) ? $data['role'] : 'no role';
