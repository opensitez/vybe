<?php
// vybe-test: php/php5_legacy/closure_this
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class Foo {
    public $x = 42;
    public function getClosure() {
        return function() { return $this->x; };
    }
}
$f = new Foo();
$fn = $f->getClosure();
