<?php
// vybe-test: php/php84/asymmetric_visibility_public_private_set
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Counter {
    public private(set) int $count = 0;
    public function increment(): void { $this->count++; }
}
$c = new Counter();
$c->increment();
$c->increment();
echo $c->count;
