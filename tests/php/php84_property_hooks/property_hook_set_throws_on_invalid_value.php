<?php
// vybe-test: php/php84_property_hooks/property_hook_set_throws_on_invalid_value
// origin: languages/php/tests/php/test_php84_property_hooks.rs

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

class PositiveInt {
    public int $value {
        set(int $v) {
            if ($v <= 0) throw new \RangeException("Must be positive");
            $this->value = $v;
        }
    }
}
$p = new PositiveInt();
try {
    $p->value = -5;
} catch (\RangeException $e) {
    echo $e->getMessage();
}

__vybe_check(ob_get_clean(), "Must be positive");
