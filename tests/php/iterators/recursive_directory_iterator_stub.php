<?php
// vybe-test: php/iterators/recursive_directory_iterator_stub
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

// RecursiveDirectoryIterator usage pattern
class TreeNode implements RecursiveIterator {
    private int $pos = 0;
    public function __construct(private array $children) {}
    public function current(): mixed  { return $this->children[$this->pos]; }
    public function key(): int        { return $this->pos; }
    public function next(): void      { $this->pos++; }
    public function rewind(): void    { $this->pos = 0; }
    public function valid(): bool     { return $this->pos < count($this->children); }
    public function hasChildren(): bool   { return is_array($this->current()); }
    public function getChildren(): static { return new static($this->current()); }
}
$tree = new TreeNode(['a', 'b', 'c']);
$items = [];
foreach ($tree as $item) { if (!is_array($item)) $items[] = $item; }
echo implode(',', $items);
