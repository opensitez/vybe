<?php
// vybe-test: php/php_oop_nullsafe_operator_chaining/test_nullsafe_method_then_property_and_arithmetic
// origin: languages/php/tests/php/test_php_oop_nullsafe_operator_chaining.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

class Counter {
    public function value(): int { return 4; }
}
class Holder {
    public ?Counter $counter;
    public function __construct(?Counter $counter = null) { $this->counter = $counter; }
}
$has = new Holder(new Counter());
$none = new Holder();
echo ($has->counter?->value() + 1);
echo '|';
echo ($none->counter?->value() + 1) ?? 'none';

__vybe_check(ob_get_clean(), "5|1");
