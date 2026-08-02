<?php
// vybe-test: php/control_flow_advanced/match_with_no_arg_bool_conditions
// origin: languages/php/tests/php/test_control_flow_advanced.rs

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

$x = 15;
echo match(true) { $x % 15 === 0 => 'FizzBuzz', $x % 3 === 0 => 'Fizz', $x % 5 === 0 => 'Buzz', default => (string)$x };

__vybe_check(ob_get_clean(), "FizzBuzz");
