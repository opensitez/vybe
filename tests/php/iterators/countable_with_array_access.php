<?php
// vybe-test: php/iterators/countable_with_array_access
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class DataSet implements ArrayAccess, Countable {
    private array $rows = [];
    public function offsetExists(mixed $k): bool  { return isset($this->rows[$k]); }
    public function offsetGet(mixed $k): mixed    { return $this->rows[$k]; }
    public function offsetSet(mixed $k, mixed $v): void { $this->rows[] = $v; }
    public function offsetUnset(mixed $k): void   { unset($this->rows[$k]); }
    public function count(): int { return count($this->rows); }
}
$ds = new DataSet();
$ds[] = ['id' => 1]; $ds[] = ['id' => 2]; $ds[] = ['id' => 3];
echo count($ds);
