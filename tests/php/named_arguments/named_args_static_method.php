<?php
// vybe-test: php/named_arguments/named_args_static_method
// origin: languages/php/tests/php/test_named_arguments.rs

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

class MathHelper {
    public static function power(float $base, int $exponent = 2): float {
        return pow($base, $exponent);
    }
}
echo MathHelper::power(base: 3.0, exponent: 3) . "\n";
echo MathHelper::power(base: 4.0) . "\n";

__vybe_check(ob_get_clean(), "27\n16");
