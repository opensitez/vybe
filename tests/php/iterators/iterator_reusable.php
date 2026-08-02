<?php
// vybe-test: php/iterators/iterator_reusable
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class Letters implements Iterator {
    private int $pos = 0;
    private array $letters = ['a', 'b', 'c'];
    public function current(): string { return $this->letters[$this->pos]; }
    public function key(): int { return $this->pos; }
    public function next(): void { $this->pos++; }
    public function rewind(): void { $this->pos = 0; }
    public function valid(): bool { return $this->pos < count($this->letters); }
}
$it = new Letters();
foreach ($it as $v) { echo $v; }
foreach ($it as $v) { echo $v; }
