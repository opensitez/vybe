<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_promoted_property_varargs_expansion
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs
// vybe-test-mode: compile

class UserList {
    public array $users;
    public function __construct(string ...$users) {
        $this->users = $users;
    }
}

$ul = new UserList("Alice", "Bob", "Charlie");
echo count($ul->users);
