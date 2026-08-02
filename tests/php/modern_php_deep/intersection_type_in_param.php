<?php
// vybe-test: php/modern_php_deep/intersection_type_in_param
// origin: languages/php/tests/php/test_modern_php_deep.rs
// vybe-test-mode: compile

interface Countable2 { public function size(): int; }
interface Iterable2  { public function items(): array; }
class Bag implements Countable2, Iterable2 {
    private array $data;
    public function __construct(array $d) { $this->data = $d; }
    public function size(): int { return count($this->data); }
    public function items(): array { return $this->data; }
}
function process(Countable2&Iterable2 $obj): string {
    return "size=" . $obj->size() . ",items=" . count($obj->items());
}
echo process(new Bag([1, 2, 3]));
