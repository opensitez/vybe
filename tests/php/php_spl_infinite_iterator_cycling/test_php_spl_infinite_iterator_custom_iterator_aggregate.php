<?php
// vybe-test: php/php_spl_infinite_iterator_cycling/test_php_spl_infinite_iterator_custom_iterator_aggregate
// origin: languages/php/tests/php/test_php_spl_infinite_iterator_cycling.rs
// vybe-test-mode: compile

class SimpleCollection implements IteratorAggregate {
    public function getIterator(): Traversable {
        return new ArrayIterator([10, 20]);
    }
}
$coll = new SimpleCollection();
$inf = new InfiniteIterator($coll->getIterator());
$inf->rewind();
echo $inf->current() === 10 ? "AGGREGATE_INFINITE_OK" : "FAIL";
