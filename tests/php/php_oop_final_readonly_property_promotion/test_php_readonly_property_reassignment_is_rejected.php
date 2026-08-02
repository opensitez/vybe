<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_readonly_property_reassignment_is_rejected
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs

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

class ImmutableCounter {
    public function __construct(public readonly int $value) {}

    public function safeSet(int $next): string {
        try {
            $this->value = $next;
            return "ok";
        } catch (Error $e) {
            return "error";
        }
    }
}

$counter = new ImmutableCounter(4);
echo $counter->value . "|" . $counter->safeSet(7) . "|" . $counter->value;

__vybe_check(ob_get_clean(), "4|error|4");
