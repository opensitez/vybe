<?php
// vybe-test: php/php84/asymmetric_visibility_public_protected_set
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Entity {
    public protected(set) string $id;
    public function __construct(string $id) { $this->id = $id; }
}
class User extends Entity {
    public function rename(string $id): void { $this->id = $id; }
}
$u = new User('user-1');
echo $u->id;
$u->rename('user-2');
echo $u->id;
