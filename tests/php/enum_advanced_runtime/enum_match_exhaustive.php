<?php
// vybe-test: php/enum_advanced_runtime/enum_match_exhaustive
// origin: languages/php/tests/php/test_enum_advanced_runtime.rs

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

enum Dir { case N; case S; }
function flip(Dir $d): string { return match ($d) { Dir::N => 'north', Dir::S => 'south' }; }
echo flip(Dir::N);

__vybe_check(ob_get_clean(), "north");
