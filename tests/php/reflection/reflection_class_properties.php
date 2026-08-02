<?php
// vybe-test: php/reflection/reflection_class_properties
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class User {
    public string $name;
    protected int $age;
    private string $password;
}
$rc = new ReflectionClass(User::class);
$props = $rc->getProperties();
echo count($props);
