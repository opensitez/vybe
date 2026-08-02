<?php
// vybe-test: php/reflection/reflection_method_is_static
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class Factory {
    public static function create(): static { return new static(); }
    public function doSomething(): void {}
}
$rc = new ReflectionClass(Factory::class);
echo $rc->getMethod('create')->isStatic() ? 'static' : 'instance';
echo $rc->getMethod('doSomething')->isStatic() ? 'static' : 'instance';
