<?php
// vybe-test: php/oop/method_chaining
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Builder {
    public $parts = [];
    public function add($part) { array_push($this->parts, $part); return $this; }
    public function build() { return implode(', ', $this->parts); }
}
$b = new Builder();
echo $b->add('a')->add('b')->add('c')->build();
