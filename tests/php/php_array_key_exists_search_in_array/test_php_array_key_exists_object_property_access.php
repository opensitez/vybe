<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_key_exists_object_property_access
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs
// vybe-test-mode: compile

class Container implements ArrayAccess {
    private array $container = ["foo" => "bar"];
    public function offsetExists(mixed $offset): bool { return isset($this->container[$offset]); }
    public function offsetGet(mixed $offset): mixed { return $this->container[$offset]; }
    public function offsetSet(mixed $offset, mixed $value): void { $this->container[$offset] = $value; }
    public function offsetUnset(mixed $offset): void { unset($this->container[$offset]); }
}
$c = new Container();
echo isset($c["foo"]) ? "OFFSET_EXISTS_OK" : "FAIL";
