<?php
// vybe-test: php/reflection/reflection_property_type
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

class TypedProps {
    public int $count = 0;
    public ?string $label = null;
}
$rc = new ReflectionClass(TypedProps::class);
$prop = $rc->getProperty('count');
echo $prop->getType()->getName();
