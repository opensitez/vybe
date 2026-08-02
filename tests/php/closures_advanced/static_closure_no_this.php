<?php
// vybe-test: php/closures_advanced/static_closure_no_this
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Foo {
    public int $x = 5;
    public function getStatic(): Closure {
        return static function() {
            // $this is not available here
            return 'static closure';
        };
    }
}
$f = new Foo();
echo $f->getStatic()();
