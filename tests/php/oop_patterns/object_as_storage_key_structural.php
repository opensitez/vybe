<?php
// vybe-test: php/oop_patterns/object_as_storage_key_structural
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Token {
    public function __construct(public string $value) {}
}
$storage = new SplObjectStorage();
$t1 = new Token('abc');
$t2 = new Token('xyz');
$storage->attach($t1, 'meta-for-abc');
$storage->attach($t2, 'meta-for-xyz');
echo $storage->count();
echo $storage[$t1];
