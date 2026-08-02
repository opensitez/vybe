<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_array_column_object_properties
// origin: languages/php/tests/php/test_php_array_fill_range_pad_column.rs
// vybe-test-mode: compile

class Person {
    public function __construct(public int $id, public string $name) {}
}
$people = [new Person(1, "Alice"), new Person(2, "Bob")];
$names = array_column($people, "name");
echo implode(",", $names);
