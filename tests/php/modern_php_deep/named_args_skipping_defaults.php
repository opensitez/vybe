<?php
// vybe-test: php/modern_php_deep/named_args_skipping_defaults
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

function create(string $type, string $color = "black", int $size = 10): string {
    return "$color $type (size $size)";
}
echo create(type: "circle", size: 20);

__vybe_check(ob_get_clean(), "black circle (size 20)");
