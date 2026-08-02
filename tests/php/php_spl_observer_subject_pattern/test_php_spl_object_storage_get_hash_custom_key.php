<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_object_storage_get_hash_custom_key
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs
// vybe-test-mode: compile

class CustomHashStorage extends SplObjectStorage {
    public function getHash(object $object): string {
        return spl_object_hash($object);
    }
}

$chs = new CustomHashStorage();
$o = new stdClass();
$chs->attach($o);
echo $chs->contains($o) ? "CUSTOM_HASH_OK" : "FAIL";
