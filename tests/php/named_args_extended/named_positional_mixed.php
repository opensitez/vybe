<?php
// vybe-test: php/named_args_extended/named_positional_mixed
// origin: languages/php/tests/php/test_named_args_extended.rs

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

function box(string $color, int $width, int $height): string {
    return "$color {$width}x$height";
}
echo box('red', height: 10, width: 5);

__vybe_check(ob_get_clean(), "red 5x10");
