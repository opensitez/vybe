<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_subject_weak_reference_storage
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs
// vybe-test-mode: compile

$storage = new SplObjectStorage();
$obj = new stdClass();
$storage->attach($obj);
echo $storage->contains($obj) ? "STORAGE_OK" : "FAIL";
