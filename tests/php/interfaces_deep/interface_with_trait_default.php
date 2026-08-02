<?php
// vybe-test: php/interfaces_deep/interface_with_trait_default
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Hashable { public function hash(): string; }
trait DefaultHash {
    public function hash(): string { return md5(serialize($this)); }
}
class User implements Hashable {
    use DefaultHash;
    public function __construct(public string $name) {}
}
$u = new User('alice');
echo strlen($u->hash()) === 32 ? 'valid hash' : 'invalid hash';
