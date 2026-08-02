<?php
// vybe-test: php/spl/spl_object_storage_iterate
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$store = new SplObjectStorage();
for ($i = 0; $i < 3; $i++) {
    $obj = new stdClass();
    $obj->id = $i;
    $store->attach($obj, "info-$i");
}
$ids = [];
$store->rewind();
while ($store->valid()) {
    $ids[] = $store->getInfo();
    $store->next();
}
sort($ids);
echo implode(',', $ids);
