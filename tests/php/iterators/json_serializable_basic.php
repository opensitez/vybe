<?php
// vybe-test: php/iterators/json_serializable_basic
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class Money implements JsonSerializable {
    public function __construct(private int $cents, private string $currency = 'USD') {}
    public function jsonSerialize(): array {
        return ['amount' => $this->cents / 100, 'currency' => $this->currency];
    }
}
$m = new Money(1999, 'EUR');
echo json_encode($m);
