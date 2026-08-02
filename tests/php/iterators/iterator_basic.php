<?php
// vybe-test: php/iterators/iterator_basic
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class NumberRange implements Iterator {
    private int $current;
    public function __construct(
        private int $start,
        private int $end
    ) { $this->current = $start; }
    public function current(): int { return $this->current; }
    public function key(): int { return $this->current - $this->start; }
    public function next(): void { $this->current++; }
    public function rewind(): void { $this->current = $this->start; }
    public function valid(): bool { return $this->current <= $this->end; }
}
$range = new NumberRange(1, 5);
foreach ($range as $k => $v) { echo "$k:$v "; }
