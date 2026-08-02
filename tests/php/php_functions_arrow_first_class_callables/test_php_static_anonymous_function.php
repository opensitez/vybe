<?php
// vybe-test: php/php_functions_arrow_first_class_callables/test_php_static_anonymous_function
// origin: languages/php/tests/php/test_php_functions_arrow_first_class_callables.rs
// vybe-test-mode: compile

class Container {
    public function getClosure() {
        return static function() {
            return "no_this_binding";
        };
    }
}

$c = new Container();
$fn = $c->getClosure();
echo $fn();
