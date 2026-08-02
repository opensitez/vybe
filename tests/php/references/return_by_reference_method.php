<?php
// vybe-test: php/references/return_by_reference_method
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

class Counter {
    private int $val = 0;
    public function &getValue(): int { return $this->val; }
}
$c = new Counter();
$ref = &$c->getValue();
$ref = 42;
echo $c->getValue();
