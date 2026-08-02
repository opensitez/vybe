<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_object_storage_remove_all_and_add_all
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs
// vybe-test-mode: compile

$s1 = new SplObjectStorage();
$s2 = new SplObjectStorage();
$o1 = new stdClass(); $o2 = new stdClass();

$s1->attach($o1);
$s1->attach($o2);
$s2->addAll($s1);

echo count($s2) === 2 ? "ADD_ALL_OK" : "FAIL";
$s2->removeAll($s1);
echo count($s2) === 0 ? " REMOVE_ALL_OK" : " FAIL";
