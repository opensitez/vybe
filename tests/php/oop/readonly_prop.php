<?php
// vybe-test: php/oop/readonly_prop
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class User { public readonly string $name; public function __construct(string $n) { $this->name = $n; } } $u = new User('Alice');
