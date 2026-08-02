<?php
// vybe-test: php/class_inspection/property_exists_class_and_object
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class User {
    public string $name;
    protected int $age = 0;
}
$u = new User();
$u->name = 'Alice';
echo property_exists($u, 'name') ? 'yes' : 'no';
echo property_exists($u, 'age') ? 'yes' : 'no';
echo property_exists($u, 'email') ? 'yes' : 'no';
