<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_object_storage_associated_data
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs
// vybe-test-mode: compile

$storage = new SplObjectStorage();
$user = new stdClass();
$storage[$user] = ["role" => "admin"];

echo $storage[$user]["role"];
