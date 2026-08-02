<?php
// vybe-test: php/operators_runtime/unary_and_nullsafe_truthiness_runtime
// origin: languages/php/tests/php/test_operators_runtime.rs

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

class Holder {
    public ?string $value = null;
}
$holder = new Holder();
echo ($holder->value ?? 'none') . '|';
echo (!$holder->value) . '|';
echo ($holder->value ?: 'fallback') . '|';
$holder->value = 'x';
echo ($holder->value ?? 'none') . '|';
echo (!$holder->value) . '|';
echo ($holder->value ?: 'fallback');

__vybe_check(ob_get_clean(), "none|1|fallback|x|0|x");
