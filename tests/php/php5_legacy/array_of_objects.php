<?php
// vybe-test: php/php5_legacy/array_of_objects
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class Item { public $name; public function __construct($n) { $this->name = $n; } }
$items = [new Item('a'), new Item('b'), new Item('c')];
foreach ($items as $item) { echo $item->name; }
