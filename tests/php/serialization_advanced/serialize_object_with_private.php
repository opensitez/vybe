<?php
// vybe-test: php/serialization_advanced/serialize_object_with_private
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

class Secret {
    private string $password;
    public function __construct(string $pw) { $this->password = $pw; }
    public function getPassword(): string { return $this->password; }
}
$s = new Secret('abc123');
$ser = serialize($s);
$s2 = unserialize($ser);
echo $s2->getPassword();
