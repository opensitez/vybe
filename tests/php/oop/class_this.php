<?php
// vybe-test: php/oop/class_this
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Counter { public $count = 0; public function inc() { $this->count++; return $this; } } $c = new Counter(); $c->inc()->inc();
