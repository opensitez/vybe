<?php
// vybe-test: php/iterators/array_access_exists_unset
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class Registry implements ArrayAccess {
    private array $store = [];
    public function offsetExists(mixed $k): bool  { return array_key_exists($k, $this->store); }
    public function offsetGet(mixed $k): mixed    { return $this->store[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->store[$k] = $v; }
    public function offsetUnset(mixed $k): void   { unset($this->store[$k]); }
}
$r = new Registry();
$r['key'] = 'value';
echo isset($r['key']) ? 'exists' : 'missing';
unset($r['key']);
echo isset($r['key']) ? 'exists' : 'missing';
