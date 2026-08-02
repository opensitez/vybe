<?php
// vybe-test: php/declare/strict_types_class_method
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
class Money {
    public function __construct(private int $cents) {}
    public function add(int $cents): static {
        $this->cents += $cents;
        return $this;
    }
    public function format(): string { return '$' . number_format($this->cents / 100, 2); }
}
$m = new Money(100);
echo $m->add(50)->format();
