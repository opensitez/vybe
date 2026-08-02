<?php
// vybe-test: php/reflection/reflection_method_invoke
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Math { public function mul(int $a, int $b): int { return $a * $b; } }
$obj = new Math();
$method = new ReflectionMethod(Math::class, 'mul');
echo $method->invoke($obj, 6, 7);
