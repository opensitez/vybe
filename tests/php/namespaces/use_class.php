<?php
// vybe-test: php/namespaces/use_class
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Models;
class User {
    public function __construct(public string $name) {}
}

namespace App;
use Models\User;
$u = new User('Alice');
echo $u->name;
