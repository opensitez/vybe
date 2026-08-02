<?php
// vybe-test: php/iterators/iterator_infinite_with_limit
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class Fibonacci implements Iterator {
    private int $a = 0, $b = 1, $step = 0;
    public function current(): int  { return $this->a; }
    public function key(): int      { return $this->step; }
    public function next(): void    { [$this->a, $this->b] = [$this->b, $this->a + $this->b]; $this->step++; }
    public function rewind(): void  { $this->a = 0; $this->b = 1; $this->step = 0; }
    public function valid(): bool   { return true; }
}
$fib = new Fibonacci();
$result = [];
$fib->rewind();
for ($i = 0; $i < 8; $i++) { $result[] = $fib->current(); $fib->next(); }
echo implode(',', $result);
