<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_object_storage_array_access_unsetting
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs
// vybe-test-mode: compile

$s = new SplObjectStorage();
$o = new stdClass();
$s[$o] = "value";
unset($s[$o]);
echo !isset($s[$o]) ? "UNSET_OK" : "FAIL";
