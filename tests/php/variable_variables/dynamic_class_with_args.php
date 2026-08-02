<?php
// vybe-test: php/variable_variables/dynamic_class_with_args
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Color {
    public function __construct(private string $name, private string $hex) {}
    public function __toString(): string { return "{$this->name}:{$this->hex}"; }
}
$className = 'Color';
$obj = new $className('red', '#FF0000');
echo $obj;
