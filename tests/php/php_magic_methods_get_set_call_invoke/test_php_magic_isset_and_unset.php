<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_isset_and_unset
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs
// vybe-test-mode: compile

class DataBag {
    private array $data = ["a" => 1];
    public function __isset(string $name): bool {
        return isset($this->data[$name]);
    }
    public function __unset(string $name): void {
        unset($this->data[$name]);
    }
}

$b = new DataBag();
echo isset($b->a) ? "YES" : "NO";
unset($b->a);
echo isset($b->a) ? "YES" : "NO";
