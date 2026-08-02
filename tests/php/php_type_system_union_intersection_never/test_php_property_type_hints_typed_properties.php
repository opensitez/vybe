<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php_property_type_hints_typed_properties
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs
// vybe-test-mode: compile

class UserProfile {
    public string $name;
    public int $age;
    public ?string $bio = null;
    public array $tags = [];
}

$up = new UserProfile();
$up->name = "Bob";
$up->age = 30;
echo "{$up->name} {$up->age}";
