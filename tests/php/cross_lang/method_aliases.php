<?php
// vybe-test: php/cross_lang/method_aliases
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

// PHP class methods are auto-aliased so other languages can call them:
// toString → __str__ (Python) → ToString (VB/C#)
// contains → __contains__ (Python) → includes (JS)
class Collection {
    public $items = [];
    public function add($item) { array_push($this->items, $item); return $this; }
    public function contains($item) { return in_array($item, $this->items); }
    public function toString() { return implode(', ', $this->items); }
    public function count() { return count($this->items); }
}
$c = new Collection();
$c->add('a')->add('b');
echo $c->toString(); // Also callable as __str__() from Python
echo $c->contains('a'); // Also callable as includes() from JS
echo $c->count(); // Also callable as __len__() from Python
