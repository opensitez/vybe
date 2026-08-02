<?php
// vybe-test: php/reflection/reflection_class_methods
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Calculator {
    public function add(int $a, int $b): int { return $a + $b; }
    public function sub(int $a, int $b): int { return $a - $b; }
    private function secret(): void {}
}
$rc = new ReflectionClass(Calculator::class);
$public = $rc->getMethods(ReflectionMethod::IS_PUBLIC);
echo count($public);
