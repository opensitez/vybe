<?php
// vybe-test: php/variable_variables/dynamic_method_call
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Greeter {
    public function hello(): string { return "hello"; }
    public function goodbye(): string { return "goodbye"; }
}
$g = new Greeter();
$method = 'hello';
echo $g->$method();
