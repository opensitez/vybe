<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_debug_info_custom_output
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs
// vybe-test-mode: compile

class SensitiveModel {
    private string $password = "secret123";
    public function __debugInfo(): array {
        return ["password" => "******"];
    }
}

$m = new SensitiveModel();
var_dump($m);
