<?php
// vybe-test: php/iterators/countable_basic
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class WordList implements Countable {
    private array $words = [];
    public function add(string $w): void { $this->words[] = $w; }
    public function count(): int { return count($this->words); }
}
$wl = new WordList();
$wl->add('hello'); $wl->add('world');
echo count($wl);
