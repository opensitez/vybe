<?php
// vybe-test: php/oop/readonly_class
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

readonly class Dto { public function __construct(public string $name, public int $age) {} } $d = new Dto('Alice', 30);
