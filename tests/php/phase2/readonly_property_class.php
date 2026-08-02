<?php
// vybe-test: php/phase2/readonly_property_class
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

class User {
    public readonly string $name;
    public function __construct(string $name) {
        $this->name = $name;
    }
}
$u = new User('Alice');
echo $u->name;
