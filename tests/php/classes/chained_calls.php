<?php
// vybe-test: php/classes/chained_calls
// origin: languages/php/tests/php/test_classes.rs
// vybe-test-mode: compile

class Builder { public $val = ''; public function add($s) { $this->val .= $s; return $this; } } $b = new Builder(); $b->add('a')->add('b');
