<?php
// vybe-test: php/iterators/iterator_aggregate_wrapped
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class FilteredCollection implements IteratorAggregate {
    public function __construct(private array $items, private callable $filter) {}
    public function getIterator(): ArrayIterator {
        return new ArrayIterator(array_values(array_filter($this->items, $this->filter)));
    }
}
$evens = new FilteredCollection([1, 2, 3, 4, 5, 6], fn($n) => $n % 2 === 0);
foreach ($evens as $v) { echo $v . ' '; }
