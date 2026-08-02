<?php
// vybe-test: php/spread_operator/argument_spread_with_defaulted_parameters
// origin: languages/php/tests/php/test_spread_operator.rs

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

function greet(string $name, string $title = 'mr.', string $suffix = ''): string {
    return trim("$title $name $suffix");
}
echo greet(...['Doe']);
echo '|';
echo greet(...['Jane', 'Ms.', 'Jr.']);

__vybe_check(ob_get_clean(), "mr. Doe|Ms. Jane Jr.");
