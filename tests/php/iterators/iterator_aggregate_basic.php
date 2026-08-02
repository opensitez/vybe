<?php
// vybe-test: php/iterators/iterator_aggregate_basic
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class Collection implements IteratorAggregate {
    private array $items = [];
    public function add(mixed $item): void { $this->items[] = $item; }
    public function getIterator(): ArrayIterator { return new ArrayIterator($this->items); }
}
$c = new Collection();
$c->add('a'); $c->add('b'); $c->add('c');
foreach ($c as $item) { echo $item; }
