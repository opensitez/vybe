<?php
// vybe-test: php/modern_php_deep/asymmetric_visibility_structural
// origin: languages/php/tests/php/test_modern_php_deep.rs
// vybe-test-mode: compile

class Counter {
    public private(set) int $count = 0;
    public function increment(): void { $this->count++; }
}
$c = new Counter();
$c->increment();
$c->increment();
echo $c->count;
