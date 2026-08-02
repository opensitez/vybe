<?php
// vybe-test: php/php_functions_arrow_first_class_callables/test_php_first_class_callable_static_method
// origin: languages/php/tests/php/test_php_functions_arrow_first_class_callables.rs
// vybe-test-mode: compile

class Utils {
    public static function format(string $str): string {
        return strtoupper($str);
    }
}

$formatter = Utils::format(...);
echo $formatter("test");
