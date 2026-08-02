<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_property_hooks_get_set_block_body
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs

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
    private int $count = 0;

    public int $value {
        get {
            return $this->count;
        }
        set {
            if ($value < 0) {
                throw new InvalidArgumentException("Count cannot be negative");
            }
            $this->count = $value;
        }
    }
}

$c = new Counter();
$c->value = 10;
echo $c->value;

__vybe_check(ob_get_clean(), "10");
