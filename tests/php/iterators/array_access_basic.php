<?php
// vybe-test: php/iterators/array_access_basic
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class TypedCollection implements ArrayAccess {
    private array $data = [];
    public function offsetExists(mixed $offset): bool { return isset($this->data[$offset]); }
    public function offsetGet(mixed $offset): mixed   { return $this->data[$offset] ?? null; }
    public function offsetSet(mixed $offset, mixed $value): void {
        if ($offset === null) { $this->data[] = $value; }
        else { $this->data[$offset] = $value; }
    }
    public function offsetUnset(mixed $offset): void { unset($this->data[$offset]); }
}
$c = new TypedCollection();
$c[] = 'first';
$c[] = 'second';
$c['named'] = 'third';
echo $c[0] . ',' . $c['named'];
