<?php
// vybe-test: php/oop/static_factory
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class User { public $name; public function __construct($n) { $this->name = $n; } public static function create($n) { return new User($n); } } $u = User::create('Alice');
