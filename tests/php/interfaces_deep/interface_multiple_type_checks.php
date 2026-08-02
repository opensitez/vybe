<?php
// vybe-test: php/interfaces_deep/interface_multiple_type_checks
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Countable2  { public function count2(): int; }
interface Iterable2   { public function toArray(): array; }
class Collection implements Countable2, Iterable2 {
    private array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function count2(): int { return count($this->items); }
    public function toArray(): array { return $this->items; }
}
$c = new Collection([1, 2, 3]);
echo ($c instanceof Countable2) ? 'countable' : 'not countable';
echo ($c instanceof Iterable2)  ? ':iterable' : ':not iterable';
echo ':' . $c->count2();
