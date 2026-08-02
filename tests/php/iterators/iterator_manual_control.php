<?php
// vybe-test: php/iterators/iterator_manual_control
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class Counter implements Iterator {
    private int $i = 0;
    public function __construct(private int $max) {}
    public function current(): int  { return $this->i; }
    public function key(): int      { return $this->i; }
    public function next(): void    { $this->i++; }
    public function rewind(): void  { $this->i = 0; }
    public function valid(): bool   { return $this->i < $this->max; }
}
$c = new Counter(3);
$c->rewind();
while ($c->valid()) {
    echo $c->current() . ' ';
    $c->next();
}
