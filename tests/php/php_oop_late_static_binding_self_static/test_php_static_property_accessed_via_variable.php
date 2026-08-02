<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_static_property_accessed_via_variable
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs
// vybe-test-mode: compile

class ConfigHolder {
    public static string $env = "staging";
}

$className = "ConfigHolder";
echo $className::$env;
