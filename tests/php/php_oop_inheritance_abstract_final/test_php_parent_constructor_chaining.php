<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_parent_constructor_chaining
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs
// vybe-test-mode: compile

class BaseEntity {
    public int $id;
    public function __construct(int $id) {
        $this->id = $id;
    }
}

class UserEntity extends BaseEntity {
    public string $email;
    public function __construct(int $id, string $email) {
        parent::__construct($id);
        $this->email = $email;
    }
}

$u = new UserEntity(1, "user@example.com");
echo "{$u->id}: {$u->email}";
